VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmGuichet 
   AutoRedraw      =   -1  'True
   Caption         =   "Guichet"
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9420
   Icon            =   "guichet.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6375
   ScaleWidth      =   9420
   Begin VB.ListBox lstDevise 
      Height          =   3765
      Left            =   120
      TabIndex        =   28
      Top             =   1800
      Width           =   2175
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   500
      Left            =   8900
      Picture         =   "guichet.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   0
      Width           =   500
   End
   Begin VB.CommandButton cmdContext 
      BackColor       =   &H00C0C0C0&
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
      Height          =   500
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   0
      Width           =   1200
   End
   Begin TabDlg.SSTab sstabGuichet 
      Height          =   4815
      Left            =   0
      TabIndex        =   6
      Top             =   1440
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   8493
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   -2147483630
      TabCaption(0)   =   "Opération"
      TabPicture(0)   =   "guichet.frx":0544
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "picGuichet"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraOpération"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraDevise2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraDevise1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fraOptions"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Liste des opérations"
      TabPicture(1)   =   "guichet.frx":0560
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fgOpération"
      Tab(1).ControlCount=   1
      Begin VB.Frame fraOptions 
         BackColor       =   &H00C0C0FF&
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
         Height          =   3735
         Left            =   5640
         TabIndex        =   35
         Top             =   360
         Width           =   3375
         Begin MSComCtl2.DTPicker txtAmjValeur 
            Height          =   300
            Left            =   2040
            TabIndex        =   36
            Top             =   360
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CalendarBackColor=   16777215
            CalendarForeColor=   0
            CalendarTitleBackColor=   8421504
            CalendarTitleForeColor=   16777215
            CalendarTrailingForeColor=   12632256
            CustomFormat    =   "dd  MM yyy"
            Format          =   22937603
            CurrentDate     =   36299
            MaxDate         =   401768
            MinDate         =   -328351
         End
         Begin MSComCtl2.DTPicker txtAmjSelect 
            Height          =   300
            Left            =   2040
            TabIndex        =   37
            Top             =   960
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
            Format          =   22937603
            CurrentDate     =   36299
            MaxDate         =   401768
            MinDate         =   -328351
         End
         Begin VB.Label lblAmjValeur 
            BackColor       =   &H00C0C0FF&
            Caption         =   "date Valeur :"
            Height          =   255
            Left            =   240
            TabIndex        =   39
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label lblAmjSelect 
            BackColor       =   &H00C0C0FF&
            Caption         =   "Sélection des opérations du"
            Height          =   375
            Left            =   240
            TabIndex        =   38
            Top             =   960
            Width           =   1215
         End
      End
      Begin VB.Frame fraDevise1 
         Caption         =   "Devise des Espèces"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2800
         Left            =   2760
         TabIndex        =   21
         Top             =   720
         Width           =   2535
         Begin VB.OptionButton optDevise1Eur 
            Caption         =   "EUR Euros"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   400
            Width           =   1700
         End
         Begin VB.OptionButton optDevise1FRF 
            Caption         =   "FRF francs français"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   800
            Width           =   1700
         End
         Begin VB.OptionButton optDevise1CHF 
            Caption         =   "CHF francs suisses"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   1200
            Width           =   1700
         End
         Begin VB.OptionButton optDevise1USD 
            Caption         =   "USD dollars"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   2000
            Width           =   1700
         End
         Begin VB.OptionButton optDevise1Autre 
            Caption         =   "Autre"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   2400
            Width           =   1935
         End
         Begin VB.OptionButton optDevise1GBP 
            Caption         =   "GBP livres sterling"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   1600
            Width           =   1700
         End
      End
      Begin VB.Frame fraDevise2 
         Caption         =   "Devise du Compte"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2800
         Left            =   6000
         TabIndex        =   14
         Top             =   720
         Width           =   2655
         Begin VB.OptionButton optDevise2Eur 
            Caption         =   "EUR Euros"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   400
            Width           =   1700
         End
         Begin VB.OptionButton optDevise2Autre 
            Caption         =   "Autre"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   2400
            Width           =   1935
         End
         Begin VB.OptionButton optDevise2USD 
            Caption         =   "USD dollars"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   2000
            Width           =   1700
         End
         Begin VB.OptionButton optDevise2GBP 
            Caption         =   "GBP livres sterling"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   1600
            Width           =   1700
         End
         Begin VB.OptionButton optDevise2FRF 
            Caption         =   "FRF francs français"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   800
            Width           =   1700
         End
         Begin VB.OptionButton optDevise2CHF 
            Caption         =   "CHF francs suisses"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   1200
            Width           =   1700
         End
      End
      Begin VB.Frame fraOpération 
         Caption         =   "Nature"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2715
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   2000
         Begin VB.OptionButton OptArbitrage 
            Caption         =   "Arbitrage"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   1200
            Width           =   1700
         End
         Begin VB.OptionButton optChange 
            Caption         =   "Change"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   1600
            Width           =   1700
         End
         Begin VB.OptionButton optVersement 
            Caption         =   "Versement espèces"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   400
            Width           =   1700
         End
         Begin VB.OptionButton optRetrait 
            Caption         =   "Retrait espèces"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   800
            Width           =   1700
         End
         Begin VB.OptionButton optRemiseChèques 
            Caption         =   "Remise de Chèques"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   2000
            Width           =   1800
         End
      End
      Begin VB.PictureBox picGuichet 
         AutoRedraw      =   -1  'True
         FontTransparent =   0   'False
         Height          =   3855
         Left            =   2400
         ScaleHeight     =   1946.154
         ScaleMode       =   0  'User
         ScaleWidth      =   3333.333
         TabIndex        =   8
         Top             =   300
         Width           =   6660
      End
      Begin MSFlexGridLib.MSFlexGrid fgOpération 
         Height          =   4290
         Left            =   -74880
         TabIndex        =   32
         Top             =   360
         Width           =   9075
         _ExtentX        =   16007
         _ExtentY        =   7567
         _Version        =   393216
         Rows            =   1
         Cols            =   12
         FixedCols       =   0
         RowHeightMin    =   350
         BackColor       =   14737632
         ForeColor       =   12582912
         ForeColorFixed  =   -2147483641
         BackColorSel    =   12648384
         BackColorBkg    =   14737632
         AllowBigSelection=   0   'False
         TextStyleFixed  =   4
         FocusRect       =   2
         HighLight       =   0
         GridLines       =   2
         FormatString    =   $"guichet.frx":057C
      End
   End
   Begin VB.CommandButton cmdInput 
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Saisie"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   3200
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   0
      Width           =   1000
   End
   Begin VB.PictureBox picCpt 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      Height          =   900
      Left            =   0
      ScaleHeight     =   840
      ScaleWidth      =   9315
      TabIndex        =   4
      Top             =   480
      Width           =   9375
   End
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6360
      TabIndex        =   3
      Top             =   0
      Width           =   2500
   End
   Begin VB.TextBox txtRecherche 
      Height          =   300
      Left            =   4920
      TabIndex        =   2
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton cmdPageNext 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   2200
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   1000
   End
   Begin VB.CommandButton cmdPagePrior 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   1000
   End
   Begin MSComCtl2.DTPicker txtAmjMax 
      Height          =   345
      Left            =   2640
      TabIndex        =   33
      Top             =   1800
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   609
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarForeColor=   0
      CalendarTitleBackColor=   8421504
      CalendarTitleForeColor=   16777215
      CalendarTrailingForeColor=   12632256
      CustomFormat    =   "dd  MM yyy"
      Format          =   22937603
      CurrentDate     =   36299
      MaxDate         =   401768
      MinDate         =   -328351
   End
   Begin MSComCtl2.DTPicker txtAmjMin 
      Height          =   420
      Left            =   5280
      TabIndex        =   34
      Top             =   1800
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarForeColor=   0
      CalendarTitleBackColor=   8421504
      CalendarTitleForeColor=   16777215
      CalendarTrailingForeColor=   12632256
      CustomFormat    =   "dd  MM yyy"
      Format          =   22937603
      CurrentDate     =   36299
      MaxDate         =   401768
      MinDate         =   -328351
   End
   Begin VB.Label lblRecherche 
      Caption         =   "Compte"
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
      Left            =   4250
      TabIndex        =   29
      Top             =   50
      Width           =   550
   End
   Begin VB.Menu mnuRecherche 
      Caption         =   "Recherche"
      Visible         =   0   'False
      Begin VB.Menu mnuRechercheCompte 
         Caption         =   "Recherche Compte"
      End
      Begin VB.Menu r 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpérationsNonValidéesUsr 
         Caption         =   "opérations non validées Utilisateur"
      End
      Begin VB.Menu mnuOpérationsNonValidées 
         Caption         =   "opérations non validées Service"
      End
      Begin VB.Menu mnuOpérationsEnAttente 
         Caption         =   "opérations en attente d'autorisation"
      End
      Begin VB.Menu mnuOpérationsNonValidéesCaisse_Print 
         Caption         =   "Liste des opérations non validées Caisse"
      End
      Begin VB.Menu mnuOpérationsNonValidéesChèque_Print 
         Caption         =   "Liste des opérations non validées Chèque"
      End
      Begin VB.Menu mnuOpérationsNonValidéesGuichet_Print 
         Caption         =   "Liste des opérations non validées Autres"
      End
      Begin VB.Menu x6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpérationsTC_Print 
         Caption         =   "Etat des 'Fiches de Change' du jour"
      End
      Begin VB.Menu mnuOpérationsTCAmj_Print 
         Caption         =   "Etat des 'Fiches de Change' autre date (cf Option)"
      End
      Begin VB.Menu r2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpérationsDuJour 
         Caption         =   "opérations du jour"
      End
      Begin VB.Menu mnuOpérationsAMJ 
         Caption         =   "opérations à une autre date (cf Option)"
      End
      Begin VB.Menu r3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuValidationDemandeCaisse 
         Caption         =   "demande de validation Caisse"
      End
      Begin VB.Menu mnuValidationDemandeChèque 
         Caption         =   "demande de validation  Chèque"
      End
      Begin VB.Menu mnuValidationDemandeJournalGuichet 
         Caption         =   "demande de validation autres opérations"
      End
      Begin VB.Menu mnuLotàValider 
         Caption         =   "Lots à valider"
      End
      Begin VB.Menu r4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCumul_Display 
         Caption         =   "Afficher les soldes Billets & Monnaies IN / OUT"
      End
      Begin VB.Menu mnuCumul_Print 
         Caption         =   "Imprimer les soldes Billets & Monnaies IN / OUT"
      End
      Begin VB.Menu mnuCumul_Print2 
         Caption         =   "Imprimer le cumul des opérations d'une période"
      End
      Begin VB.Menu x5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuContextOptions 
         Caption         =   "Options"
      End
      Begin VB.Menu mnuContextAbandonner 
         Caption         =   "Abandonner"
      End
      Begin VB.Menu mnuContextQuitter 
         Caption         =   "Quitter"
      End
   End
   Begin VB.Menu mnuLst 
      Caption         =   "Liste"
      Visible         =   0   'False
      Begin VB.Menu mnuOpérationDétail 
         Caption         =   "détail"
      End
      Begin VB.Menu mnuOpérationAvis 
         Caption         =   "duplicata avis (2ex)"
      End
      Begin VB.Menu L1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpérationAutorisationRefusée 
         Caption         =   "autorisation refusée"
      End
      Begin VB.Menu mnuOpérationAutorisationAccordée 
         Caption         =   "autorisation accordée"
      End
      Begin VB.Menu mnuOpérationAnnuler 
         Caption         =   "Annuler l'opération"
      End
      Begin VB.Menu L2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLotValider 
         Caption         =   "Valider le lot"
      End
      Begin VB.Menu mnuLotValiderAnnuler 
         Caption         =   "Annuler la  validation du lot"
      End
      Begin VB.Menu mnuValidationDemandeAnnuler 
         Caption         =   "Annuler la demande de validation"
      End
      Begin VB.Menu L 
         Caption         =   "-"
      End
      Begin VB.Menu mnuValidationDemande_Print 
         Caption         =   "réimprimer la demande de validation"
      End
      Begin VB.Menu mnuLotComptabilisé_Print 
         Caption         =   "Imprimer les écritures comptabilisées"
      End
   End
End
Attribute VB_Name = "frmGuichet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim arrTag() As Boolean, arrTagNb As Integer, lstErrClear As Boolean
Dim lastActiveControl_Name  As String, currentActiveControl_Name As String, currentAction As String
Dim MouseMoveActiveControl_Name  As String, MouseMoveActiveControl As typeUsrColor
Dim blnMsgBox_Quit As Boolean, blnAddNew As Boolean, blnGlobalControl As Boolean, blnControl As Boolean
Dim GuichetAut As typeAuthorization
Dim X As String, X1 As String, I As Integer, Msg As String, valX As String
Dim recCompte As typeCompte
Dim recGuichet As typeGuichet
Dim CV1 As typeCV, CV2 As typeCV
Dim currentAMJ As String
Dim blnDevise1 As Boolean, blnDevise2 As Boolean
Dim M_optOpération As Control, M_optDevise1 As Control, M_optOpération2 As Control, M_optDevise2 As Control
Dim mSStab_Caption As String

Dim blnfrmCompte_Show As Boolean, mfrmCompte_Show As String

Dim MmnuValidationDemandeAnnuler  As String
Dim MmnuLotValider  As String
Dim MmnuValidationDemande_Print  As String
Dim MmnuLotComptabilisé_Print  As String
Dim MmnuLotValiderAnnuler  As String


Dim fgOpération_FormatString As String, fgOpération_K As Integer
Dim fgOpération_RowDisplay As Integer, fgOpération_RowClick As Integer
Dim fgOpération_ColorClick As Long, fgOpération_ColorDisplay As Long
Dim fgOpération_Sort1 As Integer, fgOpération_Sort2 As Integer
Dim fgOpération_SortAD As Integer, fgOpération_Sort1_Old As Integer
Dim fgOpération_arrIndex As Integer

Public Sub cmdRechercher()
Me.PopupMenu mnuRecherche, vbPopupMenuRightButton

End Sub

Private Sub cmdContext_Click()
Select Case cmdContext.Caption
    Case Is = constcmdRechercher: cmdRechercher
    Case Is = constcmdAbandonner: cmdContext_Quit
End Select

End Sub

Public Sub cmdContext_Quit()
If lstDevise.Visible Then
    lstDevise.Visible = False
Else
    If fraOptions.Visible Then
        fraOptions.Visible = False
    Else
    '    If picOpération.Visible Then
    '        picOpération.Visible = False
    '    Else
            If blnMsgBox_Quit Then
                X = MsgBox("Voulez-vous réellement abandonner?", vbYesNo + vbQuestion + vbDefaultButton2, "Saisie non enregistrée")
             Else
                X = vbYes
             End If
             If X = vbYes Then Unload Me
    End If
End If

End Sub

Public Sub fgOpération_Display()
Dim I As Integer

fgOpération.Visible = True
fgOpération.Clear

fgOpération.Rows = 1
fgOpération.FormatString = fgOpération_FormatString
fgOpération.Enabled = True
For I = 1 To arrGuichetNb
    recGuichet = arrGuichet(I)
    fgOpération.Rows = fgOpération.Rows + 1
    fgOpération.Row = fgOpération.Rows - 1
    fgOpération_DisplayLine
    fgOpération.TextArray(11 + fgOpération_K) = I
Next I
If fgOpération.Rows > 1 Then fgOpération_Sort
fgOpération_Sort1 = 9: fgOpération_Sort2 = 9: fgOpération_SortAD = 6
End Sub
Public Sub fgOpération_Sort()
'fgOpération.Row = 1
'fgOpération.RowSel = fgOpération.Rows - 1

'fgOpération.Col = 9
'fgOpération.ColSel = 9
'fgOpération.Sort = 1

If fgOpération.Rows > 1 Then
    fgOpération.Row = 1
    fgOpération.RowSel = fgOpération.Rows - 1
    
    If fgOpération_Sort1_Old = fgOpération_Sort1 Then
        If fgOpération_SortAD = 5 Then
            fgOpération_SortAD = 6
        Else
            fgOpération_SortAD = 5
        End If
    Else
        fgOpération_SortAD = 5
    End If
    fgOpération_Sort1_Old = fgOpération_Sort1
    
    fgOpération.Col = fgOpération_Sort1
    fgOpération.ColSel = fgOpération_Sort2
    fgOpération.Sort = fgOpération_SortAD
End If

End Sub


Public Sub cmdContext_Return()
If sstabGuichet.Tab = 0 Then
    If cmdInput.Visible Then
        cmdInput_Click
    Else
        SendKeys "{TAB}"
    End If
End If

End Sub


Private Sub cmdContext_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdContext
End Sub

Private Sub cmdInput_Click()
If blnControl Then Opération_Control
End Sub

Private Sub cmdInput_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdInput
End Sub


Private Sub cmdPageNext_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdPageNext
End Sub


Private Sub cmdPagePrior_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdPagePrior
End Sub


Private Sub cmdPrint_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdPrint
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim X As String
If GuichetAut.Saisir Then

    mnuOpérationsNonValidéesUsr_Click
    If arrGuichetNb > 0 Then
        X = MsgBox("Vous avez des opérations en cours.Voulez-vous vraiment quitter le programme 'Guichet' ?", vbQuestion + vbYesNo, Me.Caption)
        Cancel = IIf(X = vbNo, True, False)
    End If
End If

End Sub

Private Sub fraDevise1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset
End Sub


Private Sub fraDevise2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset
End Sub


Private Sub fraOpération_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset
End Sub


Private Sub lblRecherche_Click()
mnuRechercheCompte_Click
End Sub

Private Sub lstDevise_Click()
Select Case currentActiveControl_Name
    Case "optDevise1Autre":  Opération_Devise1 mId$(lstDevise.Text, 1, 3): optDevise1Autre.Caption = CV1.DeviseLibellé
    Case "optDevise2Autre":  Opération_Devise2 mId$(lstDevise.Text, 1, 3): optDevise2Autre.Caption = CV2.DeviseLibellé
End Select
lstDevise.Visible = False
End Sub
Private Sub lstDevise_LostFocus()
lstDevise.Visible = False
End Sub

'---------------------------------------------------------
Public Sub lstDevise_Display()
'---------------------------------------------------------

lstDevise.Visible = True
Call LstDictio(889, lstDevise)
lstDevise.SetFocus

End Sub


Private Sub fgOpération_Click()
Dim X12 As String * 12

fgOpération_K = fgOpération.Row * fgOpération.Cols
If fgOpération.Row > 0 Then
    arrGuichetIndex = Val(fgOpération.TextArray(11 + fgOpération_K))

    If arrGuichetIndex > 0 And arrGuichetIndex <= arrGuichetNb Then
        mnuOpérationDétail = False
        mnuOpérationAvis = False
        mnuValidationDemandeAnnuler = False
        mnuOpérationAnnuler = False
        mnuOpérationAutorisationRefusée = False
        mnuOpérationAutorisationAccordée = False
        
        mnuLotValider = False
        mnuValidationDemande_Print = False
        mnuLotComptabilisé_Print = False
        mnuLotValiderAnnuler = False
        
        X12 = arrGuichet(arrGuichetIndex).ValidationUsr
        
        If arrGuichet(arrGuichetIndex).SaisieAmj <> DSys And mId$(X12, 1, 10) <> "$TOTAL_CTL" Then
            mnuOpérationDétail = GuichetAut.Consulter
            mnuOpérationAvis = GuichetAut.Saisir
            If Trim(arrGuichet(arrGuichetIndex).ComptaUsr) = "" Then mnuOpérationAnnuler = GuichetAut.Xspécial

        Else
       
            If mId$(X12, 1, 6) <> "$TOTAL" Then
                mnuOpérationDétail = GuichetAut.Consulter
                mnuOpérationAvis = GuichetAut.Saisir
            Else
                X = " : " & Trim(Format$(arrGuichet(arrGuichetIndex).ComptaUsr, "######0"))
                mnuValidationDemandeAnnuler.Caption = MmnuValidationDemandeAnnuler & X
                mnuLotValider.Caption = MmnuLotValider & X
                mnuLotComptabilisé_Print.Caption = MmnuLotComptabilisé_Print & X
                mnuLotValiderAnnuler.Caption = MmnuLotValiderAnnuler & X
                Select Case mId$(X12, 1, 10)
    
                    Case "$TOTAL_CTL"
                         
                                    mnuValidationDemandeAnnuler = GuichetAut.Valider
                                    mnuLotValider = GuichetAut.Valider
                                    mnuValidationDemande_Print = True
                    Case "$TOTAL_CPT"
                                    mnuLotComptabilisé_Print = True
                                    mnuLotValiderAnnuler = GuichetAut.Xspécial
                End Select
                
            End If
            
            If Trim(X12) = "" Then
                If Trim(arrGuichet(arrGuichetIndex).ComptaUsr) = "" Then
                    If Trim(arrGuichet(arrGuichetIndex).SaisieUsr) = Trim(usrId$) Then
                        mnuOpérationAnnuler = GuichetAut.Saisir
                    Else
                        mnuOpérationAnnuler = GuichetAut.Xspécial
                    End If
                End If
            Else
                If X12 = constEnAttente Then
                    mnuOpérationAutorisationRefusée = GuichetAut.Saisir
                    mnuOpérationAutorisationAccordée = GuichetAut.Saisir
                End If
            End If
        End If
        
        recGuichet = arrGuichet(arrGuichetIndex)
        
        Me.PopupMenu mnuLst, vbPopupMenuLeftButton
        'picOpération.Visible = True
        'recGuichet_Détail arrGuichet(arrGuichetIndex), picOpération
    End If
End If
End Sub



Private Sub mnuContextAbandonner_Click()
cmdContext_Quit

End Sub

Private Sub mnuContextOptions_Click()
sstabGuichet.Tab = 0
fraOptions.Visible = True
End Sub

Private Sub mnuContextQuitter_Click()
Unload Me

End Sub

Private Sub mnuCumul_Display_Click()
Dim K2 As Integer, I As Integer
Dim curDB As Currency, curCR As Currency, curX As Currency

sstabGuichet.Tab = 1: sstabGuichet.Caption = "Solde des comptes 'CAISSE'"
recGuichet_Init recGuichet
recGuichet.Method = "Cumul_Esp"
recGuichet.Société = SocId$
recGuichet.Agence = SocAgence$
recGuichet.SaisieAmj = DSys
recGuichet.ValidationAMJ = DSys

srvGuichet_Snd recGuichet

arrCumulEspèces_Init
fgOpération.Visible = True
fgOpération.Clear

fgOpération.Rows = 1
fgOpération.FormatString = fgOpération_FormatString
fgOpération.Enabled = True
For I = 1 To selCompte_Nb
    fgOpération.Rows = fgOpération.Rows + 1
    fgOpération.Row = fgOpération.Rows - 1

    fgOpération_K = (fgOpération.Row) * fgOpération.Cols
    
    Call arrCumulEspèces_Total(selCompte(I).Devise, curDB, curCR)
    G_CV1.DeviseN = selCompte(I).Devise: CV_AttributN G_CV1
 
    fgOpération.TextArray(0 + fgOpération_K) = selCompte(I).Devise
    fgOpération.TextArray(1 + fgOpération_K) = Compte_Imp(selCompte(I).Numéro)
    fgOpération.TextArray(2 + fgOpération_K) = Trim(selCompte(I).Intitulé) & "     " & G_CV1.DeviseIso
    curX = selCompte(I).SoldeInstantané + curDB + curCR
    
    K2 = IIf(curX < 0, 3, 4)
    fgOpération.TextArray(K2 + fgOpération_K) = Format(curX, "#### ### ###.00")
    fgOpération.TextArray(5 + fgOpération_K) = Format(curDB, "#### ### ###.00") & "  " & Format(curCR, "#### ### ###.00")

Next I

Call selCompte_Load(recCompte, recCompte, "End")
'$$$$$$$$$$$$$$$$$$$$$$$

End Sub

Private Sub mnuCumul_Print_Click()
recGuichet_Init recGuichet
recGuichet.Method = "Cumul_Esp"
recGuichet.Société = SocId$
recGuichet.Agence = SocAgence$
recGuichet.SaisieAmj = DSys
recGuichet.ValidationAMJ = DSys

srvGuichet_Snd recGuichet
prtGuichet_CumulX "Guichet_Cumul"
End Sub

Private Sub mnuCumul_Print2_Click()
txtAmjMin.Visible = True: txtAmjMax.Visible = True
sstabGuichet.Visible = False
End Sub

Private Sub mnuLotValiderAnnuler_Click()
recGuichet.Method = "annCompta"
srvGuichet_Update recGuichet

End Sub

Private Sub mnuOpérationAvis_Click()
G_CV1.DeviseN = recGuichet.DeviseEspèces: CV_AttributN G_CV1
G_CV2.DeviseN = recGuichet.Devise: CV_AttributN G_CV2
prtguichetX "1", recGuichet, G_CV1, G_CV2
prtguichetX "2", recGuichet, G_CV1, G_CV2

End Sub


Private Sub mnuOpérationsAMJ_Click()
Dim X As String
X = Format$(txtAmjSelect.Year, "0000") & Format$(txtAmjSelect.Month, "00") & Format$(txtAmjSelect.Day, "00")

recGuichet_Init recGuichet
recGuichet.Method = "SnapL0"
recGuichet.Société = SocId$
recGuichet.Agence = SocAgence$
recGuichet.SaisieAmj = X

arrGuichet(0) = recGuichet
arrGuichet(0).CptMvtPièce = 9999
mSStab_Caption = "Liste des opérations du " & X

Opération_Sql

End Sub

Private Sub mnuOpérationsNonValidéesCaisse_Print_Click()
Guichet_Compta.OpérationsNonvalidées_Print "Caisse"

End Sub

Private Sub mnuOpérationsNonValidéesChèque_Print_Click()
Guichet_Compta.OpérationsNonvalidées_Print "Chèque"

End Sub

Private Sub mnuOpérationsNonValidéesGuichet_Print_Click()
Guichet_Compta.OpérationsNonvalidées_Print "Guiche"

End Sub


Private Sub mnuOpérationsNonValidéesUsr_Click()
recGuichet_Init recGuichet
recGuichet.Method = "SnapL5"
recGuichet.Société = SocId$
recGuichet.Agence = SocAgence$
recGuichet.SaisieUsr = usrId
arrGuichet(0) = recGuichet
arrGuichet(0).Référence = "9999999999"

mSStab_Caption = "Liste des opérations à valider"
Opération_Sql

End Sub

Private Sub mnuOpérationsTC_Print_Click()
Guichet_Compta.OpérationsTC_Print DSys

End Sub

Private Sub mnuLotComptaAnnuler_Click()
recGuichet_Init recGuichet
recGuichet.Société = "001"
recGuichet.Agence = "001"
recGuichet.ComptaUsr = "000000000000" '"0000007762"

recGuichet.Method = "annCompta"
srvGuichet_Update recGuichet

End Sub

Private Sub mnuOpérationsTCAmj_Print_Click()
Dim X As String
X = Format$(txtAmjSelect.Year, "0000") & Format$(txtAmjSelect.Month, "00") & Format$(txtAmjSelect.Day, "00")

Guichet_Compta.OpérationsTC_Print X

End Sub

Private Sub mnuValidationDemande_Print_Click()
Me.Enabled = False
Guichet_Compta.ValidationDemande recGuichet
End Sub

Private Sub mnuValidationDemandeAnnuler_Click()
Me.Enabled = False
recGuichet.Method = "annValider"
srvGuichet_Update recGuichet
mnuLotàValider_Click
Me.Enabled = True
AppActivate Me.Caption

End Sub


Private Sub mnuOpérationAnnuler_Click()
Dim X As String

X = MsgBox("Confirmez-vous l'annulation de cette opération ?" & Chr$(13) & "Compte : " & Compte_Imp(recGuichet.Compte) & " Montant :" & Format(recGuichet.Montant, "#### ### ###.00") & "  " & recGuichet.Libellé, vbYesNo + vbQuestion + vbDefaultButton2, "Référence : " & recGuichet.Référence & " Pièce : " & Format(recGuichet.CptMvtPièce, "00000.") & Format(recGuichet.CptMvtLigne, "0000"))
If X = vbYes Then
    recGuichet.Method = constAnnulé 'constUpdate
    recGuichet.ValidationUsr = constAnnulé
    recGuichet.ValidationAMJ = DSys
    recGuichet.ValidationHMS = time_Hms
    recGuichet.ComptaUsr = constAnnulé
    srvGuichet_Update recGuichet
    arrGuichet(arrGuichetIndex) = recGuichet
    fgOpération_DisplayLine
End If

End Sub

Private Sub mnuOpérationAutorisationAccordée_Click()
recGuichet.Method = constUpdate
recGuichet.ValidationUsr = ""
srvGuichet_Update recGuichet
arrGuichet(arrGuichetIndex) = recGuichet
fgOpération_DisplayLine

End Sub

Private Sub mnuOpérationAutorisationRefusée_Click()
recGuichet.Method = constAnnulé 'constUpdate
recGuichet.ValidationUsr = constAnnulé
recGuichet.ValidationAMJ = DSys
recGuichet.ValidationHMS = time_Hms
recGuichet.ComptaUsr = constAnnulé
srvGuichet_Update recGuichet
arrGuichet(arrGuichetIndex) = recGuichet
fgOpération_DisplayLine

End Sub


Private Sub mnuLotàValider_Click()
recGuichet_Init recGuichet
recGuichet.Method = "SnapL1"
recGuichet.Société = SocId$
recGuichet.Agence = SocAgence$
recGuichet.ValidationUsr = "$TOTAL_CTL"
recGuichet.Devise = "000"
recGuichet.CptMvtPièce = 0
recGuichet.CptMvtLigne = 0

arrGuichet(0) = recGuichet
arrGuichet(0).SaisieUsr = "99999999999"
arrGuichet(0).Devise = "999"
arrGuichet(0).CodeOpération = "9999"
arrGuichet(0).CptMvtPièce = 9999999
arrGuichet(0).CptMvtLigne = 9999999

mSStab_Caption = "Liste des opérations à valider"
fgOpération_Sort1 = 1: fgOpération_Sort2 = 1: fgOpération_SortAD = 6
Opération_Sql
End Sub

Private Sub mnuOpérationDétail_Click()
Dim X As String

frmGuichetDétail.Show vbModal
frmGuichetDétail.Visible = True
'frmGuichetDétail.frmGuichetDétail_Init
X = frmGuichetDétail.Caption
AppActivate X

End Sub


Private Sub mnuOpérationsDuJour_Click()
recGuichet_Init recGuichet
recGuichet.Method = "SnapL0"
recGuichet.Société = SocId$
recGuichet.Agence = SocAgence$
recGuichet.SaisieAmj = DSys

arrGuichet(0) = recGuichet
arrGuichet(0).CptMvtPièce = 9999
mSStab_Caption = "Liste des opérations du jour"

Opération_Sql

End Sub

Private Sub mnuOpérationsEnAttente_Click()
recGuichet_Init recGuichet
recGuichet.Method = "SnapL1"
recGuichet.Société = SocId$
recGuichet.Agence = SocAgence$
recGuichet.SaisieUsr = usrId
recGuichet.ValidationUsr = constEnAttente
arrGuichet(0) = recGuichet
arrGuichet(0).SaisieUsr = "99999999999"
arrGuichet(0).Devise = "999"

mSStab_Caption = "Liste des opérations en attente"
Opération_Sql

End Sub


Private Sub mnuOpérationsNonValidées_Click()
recGuichet_Init recGuichet
recGuichet.Method = "SnapL5"
recGuichet.Société = SocId$
recGuichet.Agence = SocAgence$
arrGuichet(0) = recGuichet
arrGuichet(0).SaisieUsr = "99999999999"
arrGuichet(0).Devise = "999"

mSStab_Caption = "Liste des opérations à valider"
Opération_Sql

End Sub

Private Sub mnuValidationDemande(Msg As String)
Me.Enabled = False
mnuOpérationsEnAttente_Click
If arrGuichetNb > 0 Then
    Call MsgBox("Opérations en attente : accord ou refus", vbInformation, "Guichet : demande de validation")
    Me.Enabled = True
    AppActivate Me.Caption
    Exit Sub
End If
recGuichet_Init recGuichet
recGuichet.Method = "àValider"
recGuichet.Société = SocId$
recGuichet.Agence = SocAgence$
recGuichet.ValidationUsr = Msg
recGuichet.SaisieUsr = usrId
recGuichet.Devise = "000"
recGuichet.CptMvtPièce = 0
recGuichet.CptMvtLigne = 0
recGuichet.ValidationAMJ = DSys
recGuichet.ValidationHMS = time_Hms

arrGuichet(0) = recGuichet
arrGuichet(0).ValidationUsr = Msg
arrGuichet(0).Devise = "999"
arrGuichet(0).CodeOpération = "9999"
arrGuichet(0).CptMvtPièce = 9999999
arrGuichet(0).CptMvtLigne = 9999999

srvGuichet_UpdateKMax recGuichet, arrGuichet(0)

mSStab_Caption = constDemandeDeValidation
If recGuichet.ComptaUsr <> "0000000000" Then Guichet_Compta.ValidationDemande recGuichet
Me.Enabled = True

AppActivate Me.Caption

End Sub

Private Sub mnuLotComptabilisé_Print_Click()
Guichet_Compta.LotComptabilisé_Print recGuichet.ComptaUsr
End Sub

Private Sub mnuLotValider_Click()
Dim mGuichet As typeGuichet, X As String

Me.Enabled = False
If recGuichet.SaisieUsr = usrId Then
    Call MsgBox("Vous ne pouvez pas valider vos propres opérations.", vbCritical, "Guichet : Validation ")
Else
'End If

    X = MsgBox("Cette action est irréversible. Confirmez-vous votre demande ?", vbYesNo + vbQuestion + vbDefaultButton2, "Guichet : Validation définitive du lot " & recGuichet.ComptaUsr)
    If X = vbYes Then
    
        recGuichet.Method = "Valider"
        recGuichet.ValidationUsr = usrId
        recGuichet.ComptaAMJ = DSys
        recGuichet.ComptaHMS = time_Hms
        mGuichet = recGuichet
        srvGuichet_Update recGuichet
        Guichet_Compta.Validation mGuichet
        cmdReset
    End If
End If

Me.Enabled = True
AppActivate Me.Caption

End Sub

Private Sub mnuRechercheCompte_Click()
recCompte.Numéro = txtRecherche
frmCompte_Show
End Sub

Private Sub frmCompte_Show()
X = Space$(100)
blnfrmCompte_Show = True
mfrmCompte_Show = txtRecherche
Mid$(X, 1, 12) = "frmCompte   "
Mid$(X, 13, 12) = "frmGuichet  "
Mid$(X, 25, 10) = Space$(10)
Mid$(X, 35, 3) = "   " 'recCompte.Devise
Mid$(X, 38, 11) = recCompte.Numéro
'blnControl = False
Msg_Monitor X

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
lstErr.Height = 200
C.ForeColor = txtUsr.ForeColor
C.BackColor = txtUsr.BackColor
End Sub


'---------------------------------------------------------
Private Sub cmdPrint_Click()
'---------------------------------------------------------
'recGuichet_Display arrGuichet(1), picCpt
'X = Format$(1, "000000") & Format$(arrGuichetNb, "000000")

'prtGuichetListX X

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
End Sub

'---------------------------------------------------------
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'---------------------------------------------------------
Select Case KeyCode
    Case Is = 13: KeyCode = 0: cmdContext_Return
    Case Is = 27: cmdContext_Quit
'    Case Is = 34: cmdPageNext_Click
 '   Case Is = 33: cmdPagePrior_Click
    Case Is = 44: KeyCode = 0: frmElpPrt.prtScreen
End Select

End Sub


'---------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------
'picGuichet.Picture = LoadPicture(imgGuichet)
Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)
ReDim arrGuichet(0): arrGuichetNbMax = 0


MmnuValidationDemandeAnnuler = mnuValidationDemandeAnnuler.Caption
MmnuLotValider = mnuLotValider.Caption
MmnuLotComptabilisé_Print = mnuLotComptabilisé_Print.Caption
MmnuLotValiderAnnuler = mnuLotValiderAnnuler.Caption

recCompteInit recCompte
recCptInfoInit G_CptInfo
Call BiaPgmAut_Init("Guichet", GuichetAut)
Set M_optOpération = optRetrait
Set M_optDevise1 = optDevise1FRF
Set M_optDevise2 = optDevise2FRF
usrColor_Set
fraOptions.BackColor = dbUsr.BackColor
usrColor_Container fraOptions, dbUsr.BackColor

fgOpération_FormatString = fgOpération.FormatString
paramGuichetAMJValeur = DSys
Call DTPicker_Set(txtAmjValeur, paramGuichetAMJValeur)
Call DTPicker_Set(txtAmjMin, mId$(DSys, 1, 6) & "01")
Call DTPicker_Set(txtAmjMax, DSys)
Call DTPicker_Set(txtAmjSelect, DSys)

dbDeviseChange_Replication
If Not IsNull(Guichet_Compta.param_Init) Then Exit Sub
Call cmdReset
Guichet_Compta.CV_Reset " "

End Sub






Private Sub mnuValidationDemandeCaisse_Click()
Me.Enabled = False
mnuValidationDemande "$TOTAL-ESP"
End Sub

Private Sub mnuValidationDemandeChèque_Click()
Me.Enabled = False
mnuValidationDemande "$TOTAL-CHQ"
End Sub


Private Sub mnuValidationDemandeJournalGuichet_Click()
Me.Enabled = False
mnuValidationDemande "$TOTAL-GUI"

End Sub

Private Sub OptArbitrage_Click()
lstErr.Clear
currentAction = constArbitrage
fraDevise1.Visible = True
fraDevise1.Caption = "Compte à créditer"
fraDevise2.Enabled = True
fraDevise2.Caption = "Compte à débiter"
usrColor_Opt M_optOpération, OptArbitrage
MouseMoveActiveControl.ForeColor = OptArbitrage.ForeColor
cmdInput.Visible = GuichetAut.Saisir
If txtRecherche.Enabled Then txtRecherche.SetFocus

End Sub

Private Sub OptArbitrage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set OptArbitrage
End Sub


Private Sub optChange_Click()
lstErr.Clear

currentAction = constChange
fraDevise1.Visible = True
fraDevise1.Caption = "Versement Devise en"
fraDevise2.Enabled = True
fraDevise2.Caption = "Retrait Devise en"
usrColor_Opt M_optOpération, optChange
MouseMoveActiveControl.ForeColor = optChange.ForeColor
txtRecherche = "": picCpt.Cls: recCompteInit recCompte
cmdInput.Visible = GuichetAut.Saisir
If txtRecherche.Enabled Then txtRecherche.SetFocus

End Sub

Private Sub optChange_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optChange

End Sub


Private Sub optDevise1Autre_Click()
'optDevise1Autre = True
usrColor_Opt M_optDevise1, optDevise1Autre
MouseMoveActiveControl.ForeColor = optDevise1Autre.ForeColor
currentActiveControl_Name = optDevise1Autre.Name
lstDevise_Display

End Sub

Private Sub optDevise1Autre_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'optDevise1Autre_Click
End Sub

Private Sub optDevise1Autre_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optDevise1Autre
End Sub

Private Sub optDevise1CHF_Click()
usrColor_Opt M_optDevise1, optDevise1CHF
MouseMoveActiveControl.ForeColor = optDevise1CHF.ForeColor
Opération_Devise1 "CHF"

End Sub

Private Sub optDevise1CHF_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optDevise1CHF
End Sub


Private Sub optDevise1Eur_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optDevise1Eur
End Sub

Private Sub optDevise1FRF_Click()
usrColor_Opt M_optDevise1, optDevise1FRF
MouseMoveActiveControl.ForeColor = optDevise1FRF.ForeColor
Opération_Devise1 "FRF"
End Sub

Private Sub optDevise1FRF_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optDevise1FRF
End Sub


Private Sub optDevise1GBP_Click()
usrColor_Opt M_optDevise1, optDevise1GBP
MouseMoveActiveControl.ForeColor = optDevise1GBP.ForeColor
Opération_Devise1 "GBP"
End Sub

Private Sub optDevise1GBP_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optDevise1GBP
End Sub


Private Sub optDevise1USD_Click()
usrColor_Opt M_optDevise1, optDevise1USD
MouseMoveActiveControl.ForeColor = optDevise1USD.ForeColor
Opération_Devise1 "USD"
End Sub

Private Sub optDevise1EUR_Click()
usrColor_Opt M_optDevise1, optDevise1Eur
MouseMoveActiveControl.ForeColor = optDevise1Eur.ForeColor
Opération_Devise1 "EUR"
End Sub

Private Sub optDevise1USD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optDevise1USD
End Sub

Private Sub optDevise2Autre_Click()
'optDevise2Autre = True
usrColor_Opt M_optDevise2, optDevise2Autre
MouseMoveActiveControl.ForeColor = optDevise2Autre.ForeColor
currentActiveControl_Name = optDevise2Autre.Name
lstDevise_Display

End Sub

Private Sub optDevise2Autre_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'optDevise2Autre_Click
End Sub

Private Sub optDevise2Autre_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optDevise2Autre
End Sub

Private Sub optDevise2CHF_Click()
usrColor_Opt M_optDevise2, optDevise2CHF
MouseMoveActiveControl.ForeColor = optDevise2CHF.ForeColor
Opération_Devise2 "CHF"

End Sub

Private Sub optDevise2CHF_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optDevise2CHF
End Sub


Private Sub optDevise2Eur_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optDevise2Eur
End Sub

Private Sub optDevise2FRF_Click()
usrColor_Opt M_optDevise2, optDevise2FRF
MouseMoveActiveControl.ForeColor = optDevise2FRF.ForeColor
Opération_Devise2 "FRF"

End Sub

Private Sub optDevise2FRF_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optDevise2FRF
End Sub


Private Sub optDevise2GBP_Click()
usrColor_Opt M_optDevise2, optDevise2GBP
MouseMoveActiveControl.ForeColor = optDevise2GBP.ForeColor
Opération_Devise2 "GBP"

End Sub

Private Sub optDevise2GBP_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optDevise2GBP
End Sub


Private Sub optDevise2USD_Click()
usrColor_Opt M_optDevise2, optDevise2USD
MouseMoveActiveControl.ForeColor = optDevise2USD.ForeColor
Opération_Devise2 "USD"

End Sub

Private Sub optDevise2EUR_Click()
usrColor_Opt M_optDevise2, optDevise2Eur
MouseMoveActiveControl.ForeColor = optDevise2Eur.ForeColor
Opération_Devise2 "EUR"
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


Private Sub optDevise2USD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optDevise2USD
End Sub

Private Sub optRemiseChèques_Click()
lstErr.Clear
currentAction = constRemiseChèques
fraDevise1.Visible = True
fraDevise1.Caption = "Devise de la remise"
fraDevise2.Enabled = True
fraDevise2.Caption = "Devise du compte"

usrColor_Opt M_optOpération, optRemiseChèques
MouseMoveActiveControl.ForeColor = optRemiseChèques.ForeColor
cmdInput.Visible = GuichetAut.Saisir
If txtRecherche.Enabled Then txtRecherche.SetFocus

End Sub

Private Sub optRemiseChèques_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optRemiseChèques
End Sub


Private Sub optRetrait_Click()
lstErr.Clear
currentAction = constRetrait
fraDevise1.Visible = True
fraDevise1.Caption = "Retrait Devise en"
fraDevise2.Enabled = True
fraDevise2.Caption = "Devise du compte"
usrColor_Opt M_optOpération, optRetrait
MouseMoveActiveControl.ForeColor = optRetrait.ForeColor
cmdInput.Visible = GuichetAut.Saisir
If txtRecherche.Enabled Then txtRecherche.SetFocus
End Sub


Private Sub optRetrait_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optRetrait
End Sub


Private Sub optVersement_Click()
lstErr.Clear
currentAction = constVersement
fraDevise1.Visible = True
fraDevise1.Caption = "Versement Devise en"
fraDevise2.Enabled = True
fraDevise2.Caption = "Devise du compte"
usrColor_Opt M_optOpération, optVersement
MouseMoveActiveControl.ForeColor = optVersement.ForeColor
cmdInput.Visible = GuichetAut.Saisir
If txtRecherche.Enabled Then txtRecherche.SetFocus
End Sub


Private Sub optVersement_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optVersement
End Sub


Private Sub sstabGuichet_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set sstabGuichet
End Sub


Private Sub txtAmjValeur_Change()
txtAmjValeur_Control

End Sub

Private Sub txtAmjValeur_GotFocus()
DTPicker_GotFocus txtAmjValeur

End Sub


Private Sub txtAmjValeur_LostFocus()
DTPicker_LostFocus txtAmjValeur
txtAmjValeur_Control

End Sub

Private Sub txtAmjValeur_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
DTPicker_GotFocus txtAmjValeur

End Sub


Private Sub txtRecherche_Change()
'cmdInput.Visible = False
picCpt.Cls
End Sub

'-------------------------------------------------'
Private Sub txtAmjValeur_Control()
'-------------------------------------------------'

Dim X As String
X = Format$(txtAmjValeur.Year, "0000") & Format$(txtAmjValeur.Month, "00") & Format$(txtAmjValeur.Day, "00")
If Not IsNumeric(X) Then
    Call lstErr_AddItem(lstErr, cmdContext, "? erreur date")
    DTPicker_Now txtAmjValeur
Else
    If mId$(X, 1, 8) > DSys Then
            Call lstErr_AddItem(lstErr, cmdContext, "?  date > jour")
        DTPicker_Now txtAmjValeur
    Else
        paramGuichetAMJValeur = mId$(X, 1, 8)
    End If
End If

End Sub




'-------------------------------------------------'
Private Sub txtRecherche_GotFocus()
'-------------------------------------------------'
Call txt_GotFocus(txtRecherche)
End Sub

'---------------------------------------------------------
Private Sub txtRecherche_KeyPress(KeyAscii As Integer)
'---------------------------------------------------------
KeyAscii = convUCase(KeyAscii)
End Sub


'---------------------------------------------------------
Public Sub cmdReset()
'---------------------------------------------------------
Opération_Init
blnControl = True
lstDevise.Visible = False
fraOptions.Visible = False
'picGuichet.Visible = True
txtRecherche = ""
recCompteInit recCompte
blnDevise1 = False: blnDevise2 = False
picCpt.Cls
mSStab_Caption = "Liste des opérations"
sstabGuichet.Tab = 1: sstabGuichet.Caption = mSStab_Caption
sstabGuichet.Tab = 0
'fraOpération.Enabled = False
'FraDevise1.Visible = False
'fraDevise2.Visible = False
currentAction = ""
optRetrait = False: optVersement = False: optRemiseChèques = False: OptArbitrage = False: optChange = False
optDevise1FRF = False: optDevise1USD = False 'True: optDevise1FRF_Click
optDevise1Eur = False: optDevise1CHF = False: optDevise1GBP = False: optDevise1Autre = False
optDevise1Autre.Caption = "Autre "
optDevise2FRF = False: optDevise2USD = False 'True: optDevise2FRF_Click
optDevise2Eur = False: optDevise2CHF = False: optDevise2GBP = False: optDevise2Autre = False
optDevise2Autre.Caption = "Autre "
cmdPagePrior.Enabled = False
cmdPageNext.Enabled = False
cmdInput.Visible = False
arrTag_Set False
lstErr.Visible = False
Call lstErr_Clear(lstErr, txtRecherche, "compte client")
cmdContext.Caption = constcmdRechercher
mnuValidationDemandeCaisse = GuichetAut.Saisir
mnuValidationDemandeChèque = GuichetAut.Saisir
mnuValidationDemandeAnnuler = GuichetAut.Valider
CV1.DeviseIso = ""
CV2.DeviseIso = ""
txtAmjValeur.Enabled = GuichetAut.Xspécial
mnuCumul_Print = GuichetAut.Valider
txtAmjMin.Visible = False: txtAmjMax.Visible = False
End Sub

'---------------------------------------------------------
Public Sub arrTag_Set(ByVal B As Boolean)
'---------------------------------------------------------
For I = 0 To arrTagNb: arrTag(I) = B: Next I
End Sub


Public Sub Msg_Rcv(txtMsg As String)
'---------------------------------------------------------
Select Case UCase$(Trim(mId$(txtMsg, 13, 12)))
    Case "FRMCOMPTE"
        txtRecherche = mId$(txtMsg, 38, 11)
        mfrmCompte_Show = txtRecherche
        recCompte.Devise = mId$(txtMsg, 35, 3)
        txtCompte_Control
End Select
End Sub

Public Sub Msg_Snd(ByVal X As String)
End Sub

Private Sub txtRecherche_LostFocus()
Call txt_LostFocus(txtRecherche)
sstabGuichet.Tab = 0
If blnControl Then Opération_Control
End Sub

Private Sub txtRecherche_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then mnuRechercheCompte_Click
End Sub

Public Sub txtCompte_Control()
X = Trim(txtRecherche)
If X = "" Then Call lstErr_Clear(lstErr, txtRecherche, "? compte"):  Exit Sub
If Not IsNumeric(X) Then
    mnuRechercheCompte_Click
    Call lstErr_AddItem(lstErr, cmdContext, "? Choisir un compte : " & X)
    Exit Sub
End If
recCompte.Numéro = Format$(Val(X), "00000000000")
If recCompte.Numéro < 10000000 Then
    mnuRechercheCompte_Click
    Call lstErr_AddItem(lstErr, cmdContext, "? Choisir un compte : " & X)
    Exit Sub
End If

recCompte.Société = SocId$
    recCompte.Agence = SocAgence$
    If Not IsNumeric(recCompte.Devise) Then recCompte.Devise = "001"
    recCompte.BiaTyp = "000"
    recCompte.BiaNum = "00"
    recCompte.Method = "SeekL1"
    X = recCompte.Devise & "." & recCompte.Numéro
    If IsNull(srvCompteFind(recCompte)) Then
        recCompte_Display recCompte, picCpt
        If currentAction <> constArbitrage Then Opération_Devise2 recCompte.Devise
            '!! solde intermédiaire
''''        If recCompte.Situation = "A" Then Call lstErr_Clear(lstErr, txtRecherche, "Le compte est annulé : " & X): Exit Sub
        If recCompte.Situation = "A" Or recCompte.Situation = "E" Then Call lstErr_Clear(lstErr, txtRecherche, "Le compte est annulé : " & X): Exit Sub
'        If recCompte.TypeGA = "A" And recCompte.BiaTyp <> "001" Then Call lstErr_AddItem(lstErr, cmdContext, "? type  = 001  : " & X): Exit Sub
'        If recCompte.TypeGA = "G" Then
            G_CptInfo.Société = recCompte.Société
            G_CptInfo.Agence = recCompte.Agence
            G_CptInfo.Devise = recCompte.Devise
            G_CptInfo.Numéro = recCompte.Numéro
            If Not srvCptInfo_ServiceAutorisé(G_CptInfo, paramGuichetService) Then
                Call lstErr_AddItem(lstErr, cmdContext, "? Service non autorisé à mouvementer ce compte")
                Exit Sub
            End If
'        End If
        
        If recCompte.Situation <> " " Then
           X = MsgBox("Compte bloqué : confirmez-vous ?", vbQuestion + vbYesNo, "Guichet : Contrôle du compte ")
           If X <> vbYes Then Call lstErr_AddItem(lstErr, cmdContext, "? compte bloqué : " & X)
        End If
        
    Else
        Call lstErr_AddItem(lstErr, txtRecherche, "? compte inconnu : " & X)
'        txtRecherche.SetFocus
'       Me.PopupMenu mnuRecherche, vbPopupMenuRightButton
End If
End Sub

Public Sub mdbCompte_Load()
recCompte.Société = recGuichet.Société
recCompte.Agence = recGuichet.Agence
recCompte.Devise = recGuichet.Devise
recCompte.Numéro = recGuichet.Compte
recCompte.BiaTyp = "000"
recCompte.BiaNum = "00"
recCompte.Method = "SeekL1"
If Not IsNull(mdbCptP0_Find(recCompte)) Then Call lstErr_AddItem(lstErr, lstErr, "? compte en " & recCompte.Devise): Exit Sub

End Sub

Public Sub Opération_Init()
cmdContext.Caption = constcmdAbandonner
picGuichet.Visible = False
fraOpération.Visible = True
fraDevise1.Visible = True
fraDevise2.Visible = True
fraOpération.Enabled = True
fraDevise1.Enabled = True
fraDevise2.Enabled = True
Opération_Devise1Opt CV1.DeviseIso
Opération_Devise2Opt CV2.DeviseIso
blnDevise1 = False: blnDevise2 = False
End Sub

Public Sub Opération_Devise1(strdev As String)
If Not optDevise1Autre Then optDevise1Autre.Caption = "Autre "
blnDevise1 = True
If IsNumeric(strdev) Then
    CV1.DeviseN = strdev
Else
    CV1.DeviseIso = strdev
End If
If Not IsNull(CV_Attribut(CV1)) Then Call lstErr_AddItem(lstErr, fraDevise1, "?Devise1 inconnue"): Exit Sub

End Sub

Public Sub Opération_Devise2(strdev As String)
If Not optDevise2Autre Then optDevise2Autre.Caption = "Autre "
blnDevise2 = True
If IsNumeric(strdev) Then
    CV2.DeviseN = strdev
Else
    CV2.DeviseIso = strdev
End If
If Not IsNull(CV_Attribut(CV2)) Then Call lstErr_AddItem(lstErr, fraDevise2, "?Devise2 inconnue"): Exit Sub

End Sub

Public Sub Opération_Devise1Opt(strdev As String)
optDevise1Autre.ForeColor = lblUsr.ForeColor
Select Case strdev
    Case "001", "FRF": optDevise1FRF = True: usrColor_Opt M_optDevise1, optDevise1FRF
    Case "130", "EUR": optDevise1Eur = True: usrColor_Opt M_optDevise1, optDevise1Eur
    Case "036", "CHF": optDevise1CHF = True: usrColor_Opt M_optDevise1, optDevise1CHF
    Case "006", "GBP": optDevise1GBP = True: usrColor_Opt M_optDevise1, optDevise1GBP
    Case "400", "USD": optDevise1USD = True: usrColor_Opt M_optDevise1, optDevise1USD
    Case Else:: usrColor_Opt M_optDevise1, optDevise1Autre
End Select

End Sub

Public Sub Opération_Devise2Opt(strdev As String)
optDevise2Autre.ForeColor = lblUsr.ForeColor
Select Case strdev
    Case "001", "FRF": optDevise2FRF = True: usrColor_Opt M_optDevise2, optDevise2FRF
    Case "130", "EUR": optDevise2Eur = True: usrColor_Opt M_optDevise2, optDevise2Eur
    Case "036", "CHF": optDevise2CHF = True: usrColor_Opt M_optDevise2, optDevise2CHF
    Case "006", "GBP": optDevise2GBP = True: usrColor_Opt M_optDevise2, optDevise2GBP
    Case "400", "USD": optDevise2USD = True: usrColor_Opt M_optDevise2, optDevise2USD
    Case Else: usrColor_Opt M_optDevise2, optDevise2Autre
End Select

End Sub

Public Sub Opération_Control()

If Not Me.Enabled Then Exit Sub
lstErr.Clear
Me.Enabled = False
blnControl = False

If CV1.DeviseIso = "   " Then Call lstErr_AddItem(lstErr, cmdContext, "? Devise 1 "): GoTo ExitSub
If CV2.DeviseIso = "   " Then Call lstErr_AddItem(lstErr, cmdContext, "? Devise 2 "): GoTo ExitSub
If Trim(currentAction) = "" Then Call lstErr_AddItem(lstErr, cmdContext, "? opération "): GoTo ExitSub

If DSys > Amj20011231 Then

    Select Case currentAction
        Case constVersement, constChange, constRemiseChèques
                    If CV2.EuroIn Then Call lstErr_AddItem(lstErr, cmdContext, "? IN 31.12.2001 : " & CV2.DeviseIso): GoTo ExitSub
       Case constArbitrage, constRetrait
                    If CV1.EuroIn Then Call lstErr_AddItem(lstErr, cmdContext, "? IN 31.12.2001 : " & CV1.DeviseIso): GoTo ExitSub
                    If CV2.EuroIn Then Call lstErr_AddItem(lstErr, cmdContext, "? IN 31.12.2001 : " & CV2.DeviseIso): GoTo ExitSub
   
        End Select
Else
    Select Case currentAction
        Case constVersement, constRetrait
                    If CV1.DeviseIso = "EUR" Then Call lstErr_AddItem(lstErr, cmdContext, "? espèces 31.12.2001 : " & CV1.DeviseIso): GoTo ExitSub
       Case constChange
                    If CV1.DeviseIso = "EUR" Then Call lstErr_AddItem(lstErr, cmdContext, "? espèces 31.12.2001 : " & CV1.DeviseIso): GoTo ExitSub
                    If CV2.DeviseIso = "EUR" Then Call lstErr_AddItem(lstErr, cmdContext, "? espèces 31.12.2001 : " & CV2.DeviseIso): GoTo ExitSub
   
        End Select

End If



If currentAction = constArbitrage Then
    If CV1.DeviseIso = CV2.DeviseIso Then Call lstErr_AddItem(lstErr, cmdContext, "? Arbitrage même devise"): GoTo ExitSub
    recCompte.Devise = CV1.DeviseN
    txtCompte_Control
Else
    If currentAction = constRemiseChèques Then
        If CV1.DeviseIso = "FRF" Or CV1.DeviseIso = "EUR" Then
            If CV2.DeviseIso <> "FRF" And CV2.DeviseIso <> "EUR" Then Call lstErr_AddItem(lstErr, cmdContext, "? Compte en FRF ou EUR")
        Else
            If CV1.DeviseIso <> CV2.DeviseIso Then Call lstErr_AddItem(lstErr, cmdContext, "? Utiliser la même devise"): GoTo ExitSub
        End If
        
    End If
End If

recCompte.Devise = CV2.DeviseN
If currentAction = constChange Then
    If CV1.DeviseIso = CV2.DeviseIso Then Call lstErr_AddItem(lstErr, cmdContext, "? Change même devise"): GoTo ExitSub
    If CV2.EuroIn Then
        recCompte.Numéro = paramGuichetBillets_In
    Else
        recCompte.Numéro = paramGuichetBillets_Out
    End If
Else
    txtCompte_Control
    If currentAction = constRemiseChèques Then
        If recCompte.TypeGA = "A" Then
            If recCompte.BiaTyp <> "001" And recCompte.BiaTyp <> "003" And recCompte.BiaTyp <> "010" And recCompte.BiaTyp <> "002" And recCompte.BiaTyp <> "016" And recCompte.BiaTyp <> "851" Then Call lstErr_AddItem(lstErr, cmdContext, "? type  = 001 003 010 016 851 : " & X): GoTo ExitSub
        End If
    Else
        If currentAction = constArbitrage Then
            If recCompte.TypeGA = "A" Then
                If recCompte.BiaTyp <> "001" And recCompte.BiaTyp <> "016" And recCompte.BiaTyp <> "028" Then Call lstErr_AddItem(lstErr, cmdContext, "? type  = 001 016 028: " & X): GoTo ExitSub
            End If
        Else
'jpl 2001.09.20       If recCompte.TypeGA = "A" And recCompte.BiaTyp <> "001" Then Call lstErr_AddItem(lstErr, cmdContext, "? type  = 001  : " & X): GoTo ExitSub
            If recCompte.TypeGA = "A" Then
                If recCompte.BiaTyp <> "001" And recCompte.BiaTyp <> "003" And recCompte.BiaTyp <> "010" And recCompte.BiaTyp <> "016" Then Call lstErr_AddItem(lstErr, cmdContext, "? type  = 001 003 010 016: " & X): GoTo ExitSub
            End If
        End If
    End If
    
End If


If lstErr.ListCount > 0 Then blnControl = True: GoTo ExitSub

If currentAction = constRetrait Then
    If CV2.DeviseIso = "EUR" Then
        If mfrmCompte_Show <> txtRecherche Then blnfrmCompte_Show = False
        If Not blnfrmCompte_Show Then
            Call lstErr_AddItem(lstErr, cmdContext, "? affichage FRF et EUR")
            frmCompte_Show
            GoTo ExitSub
        End If
    End If
End If

Msg = Space$(100)
Mid$(Msg, 1, 12) = "frmG_ESPECES"
Mid$(Msg, 13, 12) = Me.Name
Mid$(Msg, 25, 12) = currentAction

'arrCompte(0) = recCompte
'arrDevise(0) = recDevise1
'arrDevise(1) = recDevise2

Mid$(Msg, 38, 3) = CV1.DeviseIso
Mid$(Msg, 41, 3) = CV2.DeviseIso
Mid$(Msg, 44, 11) = recCompte.Numéro

If currentAction = constRemiseChèques Then
    frmGuichetChèques.Form_Init Msg
Else
    frmGuichetEspèces.Form_Init Msg
End If

'If fraOpération.Enabled Then
cmdReset
cmdContext.Caption = constcmdRechercher
If txtRecherche.Enabled Then txtRecherche.SetFocus

ExitSub:

Me.Enabled = True
    
'If cmdOk.Visible And cmdOk.Enabled Then cmdOk.SetFocus
blnControl = True
'If fraOpération.Enabled Then Opération_Init: txtRecherche = "": txtRecherche.SetFocus

End Sub


Public Sub Opération_Sql()
Dim I As Integer, xSens As String, mMethod As String
'picOpération.Visible = False

mMethod = Trim(recGuichet.Method) & "+"
sstabGuichet.Tab = 1: sstabGuichet.Caption = mSStab_Caption
fgOpération.Clear
arrGuichetNb = 0: arrGuichetIndex = 0
arrGuichetSuite = True
Do Until Not arrGuichetSuite
    srvGuichet_Monitor recGuichet
    recGuichet = arrGuichet(arrGuichetNb)
    recGuichet.Method = mMethod
Loop
fgOpération_Display
End Sub


Public Sub fgOpération_DisplayLine()
Dim K2 As Integer

fgOpération_K = (fgOpération.Row) * fgOpération.Cols
If recGuichet.Compte = "99999999999" Then
    fgOpération.Col = 1: fgOpération.CellBackColor = vbCyan: fgOpération.Text = "Lot : " & Format$(recGuichet.ComptaUsr, "000000")
    fgOpération.Col = 2: fgOpération.CellBackColor = vbCyan: fgOpération.Text = recGuichet.SaisieUsr & " : " & recGuichet.Journal
Else
    mdbCompte_Load
    fgOpération.TextArray(1 + fgOpération_K) = Compte_Imp(recGuichet.Compte)
    fgOpération.TextArray(2 + fgOpération_K) = recCompte.Intitulé
End If
 
fgOpération.TextArray(0 + fgOpération_K) = recGuichet.Devise
K2 = IIf(recGuichet.Sens = "D", 3, 4)
fgOpération.TextArray(K2 + fgOpération_K) = Format(recGuichet.Montant, "#### ### ###.00")
fgOpération.TextArray(5 + fgOpération_K) = recGuichet.Libellé
fgOpération.TextArray(6 + fgOpération_K) = recGuichet.SaisieUsr
fgOpération.TextArray(7 + fgOpération_K) = recGuichet.ValidationUsr
fgOpération.TextArray(8 + fgOpération_K) = recGuichet.CodeOpération
fgOpération.TextArray(9 + fgOpération_K) = Format(recGuichet.CptMvtPièce, "00000.") & Format(recGuichet.CptMvtLigne, "0000")
fgOpération.TextArray(10 + fgOpération_K) = recGuichet.Référence

If Trim(recGuichet.ValidationUsr) = Trim(constAnnulé) Then
         fgOpération.Col = 2: fgOpération.CellForeColor = errUsr.ForeColor
         fgOpération.TextArray(2 + fgOpération_K) = "opération annulée"
End If

End Sub

