VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCDStat 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000A&
   Caption         =   "CDStat : statistiques"
   ClientHeight    =   7035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9420
   Icon            =   "CDStat.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7035
   ScaleWidth      =   9420
   Begin VB.CommandButton cmdContext 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Abandonner"
      Height          =   500
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   0
      Width           =   1200
   End
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   450
      Left            =   5520
      TabIndex        =   6
      Top             =   0
      Width           =   3375
   End
   Begin TabDlg.SSTab sstab1 
      Height          =   6495
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   11456
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Sélection"
      TabPicture(0)   =   "CDStat.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraSelect"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraPrint"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Liste "
      TabPicture(1)   =   "CDStat.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fgSelect"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "?"
      TabPicture(2)   =   "CDStat.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraCVAmj"
      Tab(2).Control(1)=   "cmdImport"
      Tab(2).ControlCount=   2
      Begin VB.Frame fraCVAmj 
         Caption         =   "Contre-Valeur EUR à la date "
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
         Left            =   -74640
         TabIndex        =   27
         Top             =   600
         Width           =   3735
         Begin VB.CheckBox chkCVAmj 
            Caption         =   "date de contre-valeur Euro (onglet 3)"
            Height          =   255
            Left            =   360
            TabIndex        =   42
            Top             =   2880
            Width           =   3015
         End
         Begin VB.OptionButton optCVAmjJ 
            Caption         =   "du"
            Height          =   255
            Left            =   240
            TabIndex        =   31
            Top             =   1560
            Width           =   615
         End
         Begin VB.OptionButton optCVAmjF 
            Caption         =   "de FIN de MOIS de comptabilisation"
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Top             =   840
            Width           =   3375
         End
         Begin VB.OptionButton optCVAmjC 
            Caption         =   "de comptabilisation"
            Height          =   255
            Left            =   240
            TabIndex        =   28
            Top             =   360
            Width           =   2895
         End
         Begin MSComCtl2.DTPicker txtCVAmj 
            Height          =   300
            Left            =   1200
            TabIndex        =   30
            Top             =   1560
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
            Format          =   28180483
            CurrentDate     =   36299
            MaxDate         =   401768
            MinDate         =   -328351
         End
      End
      Begin VB.CommandButton cmdImport 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Importer CDDOS , CDPOS"
         Height          =   495
         Left            =   -74760
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   5760
         Width           =   2895
      End
      Begin VB.Frame fraPrint 
         Caption         =   "Etat"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5775
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   3375
         Begin VB.CommandButton cmdCDStatBenef 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Statistiques par bénéficiaire"
            Height          =   500
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   43
            Top             =   3240
            Width           =   3135
         End
         Begin VB.CommandButton cmdCDStat06 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Ouvertures nettes réajustées   Banque / Pays"
            Height          =   500
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   41
            Top             =   2640
            Width           =   3135
         End
         Begin VB.CommandButton cmdCDStat05 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Chiffre d'affaire Banque / Pays"
            Height          =   500
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   2040
            Width           =   3135
         End
         Begin VB.CommandButton cmdCDStat04 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Chiffre d'affaires (graphique)"
            Height          =   500
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   1440
            Width           =   3135
         End
         Begin VB.CommandButton cmdCDStat03 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Dossiers ouverts d'une période"
            Height          =   500
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   840
            Width           =   3135
         End
         Begin VB.CommandButton cmdCDStat01 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Engagement récap(Export)"
            Height          =   500
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   240
            Width           =   3135
         End
      End
      Begin VB.Frame fraSelect 
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
         Height          =   5775
         Left            =   3600
         TabIndex        =   4
         Top             =   480
         Width           =   5535
         Begin VB.CheckBox chkPagePersonnalisée 
            Caption         =   "Personnalistaion Référence / N° page "
            Height          =   495
            Left            =   240
            TabIndex        =   51
            Top             =   4200
            Width           =   1815
         End
         Begin VB.TextBox txtPageRéférence 
            Height          =   285
            Left            =   2160
            TabIndex        =   50
            Top             =   4440
            Width           =   2415
         End
         Begin VB.TextBox txtPageNo 
            Height          =   285
            Left            =   4680
            TabIndex        =   49
            Top             =   4440
            Width           =   615
         End
         Begin VB.TextBox txtPageDestinataire 
            Height          =   285
            Left            =   2160
            TabIndex        =   48
            Top             =   3960
            Width           =   2415
         End
         Begin VB.TextBox txtPageNb 
            Height          =   285
            Left            =   4680
            TabIndex        =   46
            Text            =   "1"
            Top             =   3960
            Width           =   615
         End
         Begin VB.TextBox txtBenef 
            Height          =   285
            Left            =   2040
            TabIndex        =   45
            Top             =   3480
            Width           =   3255
         End
         Begin VB.CheckBox chkBenef 
            Caption         =   "Bénéficiaire"
            Height          =   255
            Left            =   240
            TabIndex        =   44
            Top             =   3600
            Width           =   1335
         End
         Begin VB.CheckBox chkGraphique 
            Caption         =   "Edition graphique"
            Height          =   255
            Left            =   240
            TabIndex        =   40
            Top             =   3240
            Width           =   1695
         End
         Begin VB.CheckBox chkCVEurHistorique 
            Caption         =   "CV HISTORIQUE en Euros "
            Height          =   375
            Left            =   2520
            TabIndex        =   38
            Top             =   1440
            Width           =   2535
         End
         Begin VB.TextBox txtScaleStep 
            Height          =   285
            Left            =   3720
            TabIndex        =   37
            Text            =   "50 000 000"
            Top             =   2880
            Width           =   1335
         End
         Begin VB.TextBox txtScaleMax 
            Height          =   285
            Left            =   2040
            TabIndex        =   36
            Text            =   "300 000 000 "
            Top             =   2880
            Width           =   1215
         End
         Begin VB.CheckBox chkScale 
            Caption         =   "Echelle graphique (Max / pas)"
            Height          =   375
            Left            =   240
            TabIndex        =   35
            Top             =   2760
            Width           =   1575
         End
         Begin VB.CheckBox chkCVEurVeille 
            Caption         =   "CV VEILLE en Euros "
            Height          =   255
            Left            =   240
            TabIndex        =   34
            Top             =   1560
            Width           =   2055
         End
         Begin VB.CheckBox chkNotifié 
            Caption         =   "Notifié"
            Height          =   255
            Left            =   4200
            TabIndex        =   33
            Top             =   840
            Width           =   855
         End
         Begin VB.CheckBox chkConfirmé 
            Caption         =   "Confirmé"
            Height          =   255
            Left            =   4200
            TabIndex        =   32
            Top             =   1140
            Width           =   975
         End
         Begin VB.CheckBox chkAmj 
            Caption         =   "période"
            Height          =   255
            Left            =   240
            TabIndex        =   25
            Top             =   2400
            Width           =   1095
         End
         Begin VB.CheckBox chkCDI 
            Caption         =   "CDI"
            Height          =   255
            Left            =   4200
            TabIndex        =   20
            Top             =   540
            Width           =   735
         End
         Begin VB.CheckBox chkCDE 
            Caption         =   "CDE"
            Height          =   255
            Left            =   4200
            TabIndex        =   19
            Top             =   240
            Width           =   855
         End
         Begin VB.CheckBox chkPays 
            Caption         =   "Pays"
            Height          =   255
            Left            =   240
            TabIndex        =   17
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox txtPays 
            Height          =   285
            Left            =   1560
            TabIndex        =   16
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox txtDevise 
            Height          =   285
            Left            =   1560
            TabIndex        =   15
            Top             =   1200
            Width           =   615
         End
         Begin VB.CheckBox chkDevise 
            Caption         =   "Devise"
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   1200
            Width           =   1095
         End
         Begin VB.TextBox txtDossierMax 
            ForeColor       =   &H00C0FFC0&
            Height          =   285
            Left            =   3720
            TabIndex        =   13
            Top             =   1920
            Width           =   1335
         End
         Begin VB.TextBox txtDossierMin 
            Height          =   285
            Left            =   2040
            TabIndex        =   12
            Top             =   1920
            Width           =   1215
         End
         Begin VB.CheckBox chkDossier 
            Caption         =   "Dossier"
            Height          =   255
            Left            =   240
            TabIndex        =   11
            Top             =   1920
            Width           =   975
         End
         Begin VB.CheckBox chkCompte 
            Caption         =   "Racine"
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   720
            Width           =   1335
         End
         Begin VB.CommandButton cmdSelect 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Lancer le traitement"
            Height          =   855
            Left            =   3360
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   4800
            Width           =   1935
         End
         Begin VB.TextBox txtCompte 
            Height          =   285
            Left            =   1560
            TabIndex        =   5
            Top             =   720
            Width           =   1215
         End
         Begin MSComCtl2.DTPicker txtAmjMin 
            Height          =   300
            Left            =   2040
            TabIndex        =   23
            Top             =   2400
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
            Format          =   28180483
            CurrentDate     =   36299
            MaxDate         =   401768
            MinDate         =   -328351
         End
         Begin MSComCtl2.DTPicker txtAmjMax 
            Height          =   300
            Left            =   3720
            TabIndex        =   24
            Top             =   2400
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   529
            _Version        =   393216
            CalendarBackColor=   16777215
            CalendarForeColor=   0
            CalendarTitleBackColor=   8421504
            CalendarTitleForeColor=   16777215
            CalendarTrailingForeColor=   12632256
            CustomFormat    =   "dd  MM yyy"
            Format          =   28180483
            CurrentDate     =   36299
            MaxDate         =   401768
            MinDate         =   -328351
         End
         Begin VB.Label lblPageNb 
            Caption         =   "destinataire / nb ex"
            Height          =   375
            Left            =   480
            TabIndex        =   47
            Top             =   3960
            Width           =   1455
         End
      End
      Begin MSFlexGridLib.MSFlexGrid fgSelect 
         Height          =   5730
         Left            =   -74760
         TabIndex        =   3
         Top             =   480
         Width           =   8835
         _ExtentX        =   15584
         _ExtentY        =   10107
         _Version        =   393216
         Rows            =   1
         Cols            =   14
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
         FormatString    =   $"CDStat.frx":0496
      End
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   500
      Left            =   8880
      Picture         =   "CDStat.frx":055C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   500
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
      TabIndex        =   1
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
   End
   Begin VB.Menu mnucmdPrint 
      Caption         =   "Print"
      Visible         =   0   'False
      Begin VB.Menu mnucmdPrint_fgSelect 
         Caption         =   "Imprimer la liste"
      End
   End
End
Attribute VB_Name = "frmCDStat"
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
Dim x As String, I As Integer, Msg As String, valX As String, X1 As String, V As Variant, curX As Currency, dblX As Double, lngNb As Long
Dim CDStatAut As typeAuthorization

Dim fgSelect_FormatString As String, fgSelect_K As Integer
Dim fgSelect_RowDisplay As Integer, fgSelect_RowClick As Integer
Dim fgSelect_ColorClick As Long, fgSelect_ColorDisplay As Long
Dim fgSelect_Sort1 As Integer, fgSelect_Sort2 As Integer
Dim fgSelect_SortAD As Integer, fgSelect_Sort1_Old As Integer
Dim fgSelect_arrIndex As Integer

Dim meCDPOSPC() As typeCDPOSPC
Dim meCDPOSPC_Nb As Integer, meCDPOSPC_Index As Integer, meCDPOSPC_NbMax As Integer

Dim blnfgSelect_DisplayLine As Boolean

Dim blnSetfocus As Boolean, blnImport As Boolean

Dim paramCDSTAT_CDDOSPC As String, paramCDSTAT_CDPOSPC As String
Dim meCDStat_Prt As typeCDStat_Prt

Public Function param_Init()
Dim V
param_Init = Null

recElpTable_Init recElpTable
recElpTable.Method = "Seek="
recElpTable.Id = "CD"

recElpTable.K1 = "CDSTAT"
recElpTable.K2 = "CDDOSPC"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramCDSTAT_CDDOSPC = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmCDStat.lstErr, frmCDStat.cmdContext, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

recElpTable.K1 = "CDSTAT"
recElpTable.K2 = "CDPOSPC"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramCDSTAT_CDPOSPC = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmCDStat.lstErr, frmCDStat.cmdContext, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

Exit Function

Table_Error:
param_Init = V
Exit Function

Memo_Error:
param_Init = "Memo"
MsgBox recElpTable.Id & " : " & recElpTable.K1 & " : " & recElpTable.K2 & " : Mémo absent", vbCritical, "srvTI.Param_Init"
Exit Function

Num_Error:
param_Init = "Num"
MsgBox recElpTable.Id & " : " & recElpTable.K1 & " : " & recElpTable.K2 & " : " & recElpTable.Memo & " :Mémo non numérique", vbCritical, "srvTI.Param_Init"
End Function

Private Sub chkAmj_Click()
If chkAmj = "1" Then
    txtAmjMin.Visible = True: txtAmjMax.Visible = True: If blnSetfocus Then txtAmjMin.SetFocus
Else
    txtAmjMin.Visible = False: txtAmjMax.Visible = False
End If

End Sub

Private Sub chkBenef_Click()

If chkBenef = "1" Then
    txtBenef.Visible = True: If blnSetfocus Then txtBenef.SetFocus
Else
    txtBenef.Visible = False
End If

End Sub

Private Sub chkCompte_Click()
If chkCompte = "1" Then
    txtCompte.Visible = True: If blnSetfocus Then txtCompte.SetFocus
Else
    txtCompte.Visible = False
End If

End Sub

Private Sub chkCompte_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Set chkCompte

End Sub

Private Sub chkDevise_Click()
If chkDevise = "1" Then
    txtDevise.Visible = True: If blnSetfocus Then txtDevise.SetFocus
Else
    txtDevise.Visible = False
End If

End Sub


Private Sub chkBenef_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Set chkBenef

End Sub

Private Sub chkDevise_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Set chkDevise

End Sub

Private Sub chkDossier_Click()
If chkDossier = "1" Then
    txtDossierMin.Visible = True: txtDossierMax.Visible = True
    If blnSetfocus Then txtDossierMin.SetFocus
Else
    txtDossierMin.Visible = False: txtDossierMax.Visible = False
End If

End Sub

Private Sub chkDossier_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Set chkDossier

End Sub

Private Sub chkPagePersonnalisée_Click()
If chkPagePersonnalisée = "1" Then
    txtPageNo.Visible = True
    txtPageRéférence.Visible = True: If blnSetfocus Then txtPageRéférence.SetFocus
Else
    txtPageRéférence.Visible = False
    txtPageNo.Visible = False
End If

End Sub

Private Sub chkPays_Click()
If chkPays = "1" Then
    txtPays.Visible = True: If blnSetfocus Then txtPays.SetFocus
Else
    txtPays.Visible = False
End If

End Sub


Private Sub chkPays_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Set chkPays

End Sub

Private Sub chkScale_Click()
If chkScale = "1" Then
    txtScaleMax.Visible = True: txtScaleStep.Visible = True
    If blnSetfocus And fraSelect.Enabled Then txtScaleMax.SetFocus
Else
    txtScaleMax.Visible = False: txtScaleStep.Visible = False
End If

End Sub

Private Sub cmdCDStat04_Click()
Dim x As String
currentAction = "CDStat04"
fraSelect_Enabled
chkAmj.Value = "1"
chkDevise.Enabled = True
chkPays.Enabled = True
chkCompte.Enabled = True
chkConfirmé.Enabled = True
chkNotifié.Enabled = True
chkCVEurVeille.Enabled = True
chkCVEurHistorique.Enabled = True
chkScale.Enabled = True
chkGraphique.Enabled = True
x = mId$(DSys, 1, 4) - 1 & "0101"
Call DTPicker_Set(txtAmjMin, x)
Call DTPicker_Set(txtAmjMax, DSys)
chkPagePersonnalisée.Enabled = True
End Sub

Private Sub cmdCDStat05_Click()
Dim x As String
currentAction = "CDStat05"
fraSelect_Enabled
chkAmj.Value = "1"
chkDevise.Enabled = True
chkPays.Enabled = False
chkCompte.Enabled = False
chkConfirmé.Enabled = True
chkNotifié.Enabled = True
chkCVEurVeille.Enabled = True
chkCVEurHistorique.Enabled = True
chkScale.Enabled = False
chkGraphique.Enabled = False

x = mId$(DSys, 1, 4) - 1 & "0101"
Call DTPicker_Set(txtAmjMin, x)
Call DTPicker_Set(txtAmjMax, DSys)

End Sub

Private Sub cmdCDStat06_Click()

currentAction = "CDStat06"
fraSelect_Enabled
chkAmj.Value = "1"
chkDevise.Enabled = True
chkPays.Enabled = False
chkCompte.Enabled = False
chkConfirmé.Enabled = True
chkNotifié.Enabled = True
chkCVEurVeille.Enabled = True
chkCVEurHistorique.Enabled = True
chkScale.Enabled = False
chkGraphique.Enabled = False

x = mId$(DSys, 1, 4) - 1 & "0101"
Call DTPicker_Set(txtAmjMin, x)
Call DTPicker_Set(txtAmjMax, DSys)

End Sub

Private Sub cmdCDStatBenef_Click()

currentAction = "CDStatBenef"
fraSelect_Enabled
chkCVEurVeille.Enabled = True
chkBenef.Enabled = True
chkBenef.Value = "1"

End Sub

Private Sub cmdImport_Click()
Me.MousePointer = vbHourglass
Me.Enabled = False

cmdImport_Monitor

Me.MousePointer = 0
Me.Enabled = True

End Sub

Private Sub cmdCDStat01_Click()

currentAction = "CDStat01"
fraSelect_Enabled
cmdSelect_Click

End Sub

Private Sub CmdCDStat03_Click()

currentAction = "CDStat03"
fraSelect_Enabled
chkAmj.Value = "1"
Call DTPicker_Set(txtAmjMax, DSys)
Call DTPicker_Set(txtAmjMin, mId$(DSys, 1, 6) & "01")

End Sub

Private Sub fraPrint_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Reset
End Sub


'---------------------------------------------------------
Public Sub arrTag_Set(ByVal B As Boolean)
'---------------------------------------------------------
For I = 0 To arrTagNb: arrTag(I) = B: Next I
End Sub

Public Sub cmdContext_Quit()
blnControl = False
lstErr.Clear
If sstab1.Tab <> 0 Then
        sstab1.Tab = 0
Else
    If fraSelect.Enabled Then
        fraSelect.Enabled = False
        fraPrint.Enabled = True
    Else
        
        If currentAction = "" Then
            If blnMsgBox_Quit Then
                x = MsgBox("Voulez-vous réellement abandonner?", vbYesNo + vbQuestion + vbDefaultButton2, "Saisie non enregistrée")
            Else
               x = vbYes
            End If
            If x = vbYes Then Unload Me
        Else
            cmdReset
        End If
    End If

End If
End Sub

Public Sub cmdControl()
Dim lngX As Long, wCD As String


If Not Me.Enabled Then Exit Sub
Me.Enabled = False

blnControl = False

lstErr.Clear
lstErr.Height = 200
If lstErr.ListCount > 0 Then
    lstErr.Visible = True
End If
meCDStat_Prt.DossierMin = "": meCDStat_Prt.DossierMax = "": meCDStat_Prt.optDossier = False
meCDStat_Prt.AmjMin = "": meCDStat_Prt.AmjMax = "": meCDStat_Prt.optAmj = False
meCDStat_Prt.Devise = "": meCDStat_Prt.optDevise = False
meCDStat_Prt.Pays = "": meCDStat_Prt.optPays = False
meCDStat_Prt.Compte = "": meCDStat_Prt.optCompte = False
meCDStat_Prt.curScaleMax = 300000000
meCDStat_Prt.curScaleStep = 50000000
meCDStat_Prt.CVDevise = "EUR"
meCDStat_Prt.optConfirmé = False
meCDStat_Prt.optNotifié = False
meCDStat_Prt.SCCENR = " "
meCDStat_Prt.Etat = currentAction
meCDStat_Prt.PagePersonnalisée = False

If chkCDI = "1" Then
    meCDStat_Prt.optCDI = True: wCD = "CDI"
Else
    meCDStat_Prt.optCDI = True: wCD = ""
End If

If chkCDE = "1" Then
    meCDStat_Prt.optCDE = True: wCD = "CDE"
Else
    meCDStat_Prt.optCDE = True: wCD = ""
End If

If chkDossier = "1" Then
    meCDStat_Prt.optDossier = True
    lngX = CLng(Val(txtDossierMin))
    If lngX = 0 Then Call lstErr_AddItem(lstErr, cmdContext, "? préciser le numéro de dossier")
    meCDStat_Prt.DossierMin = wCD & Format(lngX, "000000")
    If Trim(txtDossierMax) = "" Then txtDossierMax = txtDossierMin
    lngX = CLng(Val(txtDossierMax))
    meCDStat_Prt.DossierMax = wCD & Format(lngX, "000000")
    If meCDStat_Prt.DossierMin > meCDStat_Prt.DossierMax Then Call lstErr_AddItem(lstErr, cmdContext, "? dossier Min > Max")
End If

If chkAmj = "1" Then
    meCDStat_Prt.optAmj = True
    Call DTPicker_Control(txtAmjMin, meCDStat_Prt.AmjMin)
    Call DTPicker_Control(txtAmjMax, meCDStat_Prt.AmjMax)
    If meCDStat_Prt.AmjMin > meCDStat_Prt.AmjMax Then Call lstErr_AddItem(lstErr, cmdContext, "? Amj Min > Max")
End If

If chkCompte = "1" Then
    meCDStat_Prt.optCompte = True
    meCDStat_Prt.Compte = Trim(txtCompte)
End If

If chkDevise = "1" Then
    meCDStat_Prt.optDevise = True
    meCDStat_Prt.Devise = Trim(txtDevise)
End If

If chkPays = "1" Then
    meCDStat_Prt.optPays = True
    meCDStat_Prt.Pays = Trim(txtPays)
End If

If chkCVAmj = "1" Then
    If optCVAmjF Then meCDStat_Prt.caseCVAmj = "F": meCDStat_Prt.CVAmj = "00000000"
    If optCVAmjC Then meCDStat_Prt.caseCVAmj = "C": meCDStat_Prt.CVAmj = "00000000"
    If optCVAmjJ Then meCDStat_Prt.caseCVAmj = "J": Call DTPicker_Control(txtCVAmj, meCDStat_Prt.CVAmj)
Else
    meCDStat_Prt.caseCVAmj = "J"
    Select Case currentAction
        Case "CDStat03": meCDStat_Prt.CVAmj = meCDStat_Prt.AmjMax
        Case Else: meCDStat_Prt.CVAmj = DSys
    End Select
End If

If chkGraphique = "1" Then meCDStat_Prt.optGraphique = True
If chkGraphique = "0" Then meCDStat_Prt.optGraphique = False
If chkScale = "1" Then
    meCDStat_Prt.curScaleMax = CCur(Val(txtScaleMax))
    meCDStat_Prt.curScaleStep = CCur(Val(txtScaleStep))
End If
If chkGraphique = "1" And chkScale = "0" Then Call lstErr_AddItem(lstErr, cmdContext, "? Graphique donc choix échelle")

If chkConfirmé = "1" Then meCDStat_Prt.optConfirmé = True
If chkNotifié = "1" Then meCDStat_Prt.optNotifié = True
If Not meCDStat_Prt.optConfirmé And Not meCDStat_Prt.optNotifié Then Call lstErr_AddItem(lstErr, cmdContext, "? préciser Confirmé / Notifié")

If chkCVEurVeille = "1" Then meCDStat_Prt.SCCENR = " "
If chkCVEurHistorique = "1" Then meCDStat_Prt.SCCENR = "M"
If chkCVEurVeille = "1" And chkCVEurHistorique = "1" Then Call lstErr_AddItem(lstErr, cmdContext, "? Choisir CV Veille ou Historique")

If chkCVEurVeille = "0" And chkCVEurHistorique = "0" Then
    If Trim(meCDStat_Prt.Devise) = "" Then
        Call lstErr_AddItem(lstErr, cmdContext, "? préciser DEV (non CV EUR)")
    Else
        meCDStat_Prt.CVDevise = meCDStat_Prt.Devise
    End If
End If

If chkBenef = "1" Then
    meCDStat_Prt.optBenef = True
    meCDStat_Prt.Benef = Trim(txtBenef)
End If

If chkPagePersonnalisée = "1" Then meCDStat_Prt.PagePersonnalisée = True
meCDStat_Prt.PageNb = Val(txtPageNb)
If meCDStat_Prt.PageNb = 0 Then meCDStat_Prt.PageNb = 1
If meCDStat_Prt.PageNb > 25 Then Call lstErr_AddItem(lstErr, cmdContext, "? 25 exemplaires max")
meCDStat_Prt.PageNo = Val(txtPageNo)
meCDStat_Prt.PageDestinataire = Trim(txtPageDestinataire)
meCDStat_Prt.PageRéférence = Trim(txtPageRéférence)

Me.Enabled = True

blnControl = True


End Sub


Private Sub cmdContext_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Set cmdContext

End Sub

Private Sub cmdPrint_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Set cmdPrint

End Sub

Private Sub cmdSelect_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Set cmdSelect

End Sub

Private Sub fraSelect_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub txtBenef_GotFocus()
txt_GotFocus txtBenef

End Sub

Private Sub txtBenef_LostFocus()
txt_LostFocus txtBenef

End Sub

Private Sub txtDevise_GotFocus()
txt_GotFocus txtDevise

End Sub

Private Sub txtDevise_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub

Private Sub txtBenef_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub

Private Sub txtDevise_LostFocus()
txt_LostFocus txtDevise

End Sub

Private Sub txtDossierMax_GotFocus()
txt_GotFocus txtDossierMax

End Sub

Private Sub txtDossierMax_KeyPress(KeyAscii As Integer)
Call num_KeyAscii(KeyAscii)

End Sub


Private Sub txtDossierMax_LostFocus()
txt_LostFocus txtDossierMax

End Sub


Private Sub txtdossiermin_GotFocus()

txt_GotFocus txtDossierMin

End Sub


Private Sub txtdossiermin_KeyPress(KeyAscii As Integer)
Call num_KeyAscii(KeyAscii)
End Sub


Private Sub txtdossiermin_LostFocus()
txt_LostFocus txtDossierMin
End Sub

Private Sub txtCompte_GotFocus()
txt_GotFocus txtCompte
End Sub

Private Sub txtCompte_KeyPress(KeyAscii As Integer)
Call num_KeyAscii(KeyAscii)

End Sub

Private Sub txtCompte_LostFocus()
txt_LostFocus txtCompte
End Sub

'---------------------------------------------------------
Public Sub cmdReset()
'---------------------------------------------------------
blnControl = False

usrColor_Set
currentAction = ""
cmdContext.Caption = constcmdRechercher: blnMsgBox_Quit = False
arrTag_Set False
lstErr.Visible = False
fgSelect_Reset

sstab1.Enabled = blnImport
fraPrint.Enabled = True

fraSelect.Enabled = False
chkCompte = "0": txtCompte.Visible = False
chkDossier = "0": txtDossierMin.Visible = False: txtDossierMax.Visible = False
chkDevise = "0": txtDevise.Visible = False
chkPays = "0": txtPays.Visible = False
chkAmj = "0": txtAmjMin.Visible = False: txtAmjMax.Visible = False

optCVAmjJ = True: DTPicker_Now txtCVAmj
optCVAmjC.Enabled = False: optCVAmjF.Enabled = False
chkPagePersonnalisée = "0":    txtPageRéférence.Visible = False:    txtPageNo.Visible = False


blnControl = True
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

sstab1.Tab = 1

fgSelect.Visible = True
fgSelect.Clear: fgSelect.Rows = 1: fgSelect_RowDisplay = 0

fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
fgSelect.Enabled = True
For meCDPOSPC_Index = 1 To meCDPOSPC_Nb
'    If meCDPOSPC.Method <> constIgnore And meCDPOSPC.Method <> constDelete Then
        fgSelect.Rows = fgSelect.Rows + 1
        fgSelect.Row = fgSelect.Rows - 1
        fgSelect_DisplayLine
'    End If
Next meCDPOSPC_Index

fgSelect_SortAD = 5
fgSelect_Sort1_Old = 1: fgSelect_Sort1 = 1
If fgSelect.Rows > 1 Then fgSelect_SortX 1

End Sub
Public Sub fgSelect_DisplayLine()

'xCDStat = meCDPOSPC(meCDPOSPC_Index)
fgSelect.Col = 0:
fgSelect.Col = fgSelect_arrIndex - 1: fgSelect.Text = ""
fgSelect.Col = fgSelect_arrIndex: fgSelect.Text = meCDPOSPC_Index

End Sub
Public Sub fgSelect_Load()
Dim x As String, mMethod As String


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
Dim I As Integer, x As String
For I = 1 To fgSelect.Rows - 1
    fgSelect.Row = I
    fgSelect.Col = fgSelect_arrIndex
    meCDPOSPC_Index = Val(fgSelect.Text)
    fgSelect.Col = fgSelect_arrIndex - 1
    Select Case lK
 '       Case 1: fgSelect.Text = Format$(xCDStat.EARIdRef, "00000000")
 '       Case 2: fgSelect.Text = Format$(xCDStat.MONDEV, "000000000000000.00")
        Case fgSelect_arrIndex: fgSelect.Text = Format$(meCDPOSPC_Index, "0000000000")
    End Select
Next I

fgSelect_Sort1 = fgSelect_arrIndex - 1: fgSelect_Sort2 = fgSelect_arrIndex - 1
fgSelect_Sort
End Sub


Public Sub Form_Init()
Call lstErr_Clear(lstErr, cmdContext, "Initialisation ")
DoEvents

sstab1.Tab = 0
ReDim meCDPOSPC(10)

blnControl = False
fgSelect_FormatString = fgSelect.FormatString
txtPageNb = 1
txtPageNo = ""
txtPageDestinataire = usrName
txtPageRéférence = ""
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


Private Sub cmdContext_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Select Case cmdContext.Caption
    Case Is = constcmdRechercher: Me.PopupMenu mnuContext, vbPopupMenuLeftButton
    Case Is = constcmdAbandonner: cmdContext_Quit
End Select

End Sub

Private Sub cmdPrint_Click()
Me.PopupMenu mnucmdPrint, vbPopupMenuLeftButton


End Sub

Private Sub cmdSelect_Click()
Dim I As Integer

cmdControl
If lstErr.ListCount = 0 Then

    Me.MousePointer = vbHourglass
    Me.Enabled = False
    For I = 1 To meCDStat_Prt.PageNb
        Call lstErr_Clear(lstErr, cmdPrint, I & " / " & meCDStat_Prt.PageNb & " " & currentAction): DoEvents
       Select Case currentAction
            Case "CDStat01": prtCDStat_01 meCDStat_Prt
            Case "CDStat03": prtCDStat_03 meCDStat_Prt
            Case "CDStat04": prtCDStat_04 meCDStat_Prt
            Case "CDStat05": prtCDStat_04 meCDStat_Prt  '!!!! 04 et non 05
            Case "CDStat06": prtCDStat_04 meCDStat_Prt  '!!!! 04 et non 06
            Case "CDStatBenef": prtCDStat_Benef meCDStat_Prt
        End Select
    Next I
    Me.MousePointer = 0
    Call lstErr_AddItem(lstErr, cmdPrint, "Traitement Terminé"): DoEvents
    
    Me.Enabled = True
    cmdContext_Quit
End If

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

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Reset
End Sub


Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim xStatut As String

If Y <= fgSelect.RowHeightMin Then
'    Select Case fgSelect.Col
'        Case 0: fgSelect_Sort1 = 0: fgSelect_Sort2 = 1: fgSelect_Sort
'        Case 1: fgSelect_SortX 1
'        Case 2:  fgSelect_SortX 2
'        Case 3: fgSelect_Sort1 = 3: fgSelect_Sort2 = 3: fgSelect_Sort
'        Case fgSelect_arrIndex:  fgSelect_SortX fgSelect_arrIndex
'    End Select
Else
    If fgSelect.Rows > 1 Then
        fgSelect.Col = fgSelect_arrIndex
        meCDPOSPC_Index = Val(fgSelect.Text)
       ' mCDStat = meCDPOSPC(meCDPOSPC_Index)
       ' xCDStat = mCDStat
        Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
        
       ' fraOpération_Display
   
       If Button = vbRightButton Then
      '      mnufgSelect_Display = CDStatAut.Consulter
            Me.PopupMenu mnufgSelect, vbPopupMenuLeftButton
       Else
        '    If CDStatAut.Consulter Then mnufgSelect_Update_Click
       
       End If
    End If
End If

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
Me.MousePointer = vbHourglass
Me.Enabled = False
mWindowState = Me.WindowState
If Me.WindowState <> frmUsr_Windowstate Then Me.WindowState = frmUsr_Windowstate

Call BiaPgmAut_Init(mId$(Msg, 1, 12), CDStatAut)

blnSetfocus = True

blnImport = True
param_Init

'Test si les fichiers sont déjà chargés

tableCDDOSPX_Open
recCDDOSPX_Init recCDDOSPX
recCDDOSPX.Method = "MoveFirst"
intReturn = tableCDDOSPX_Read(recCDDOSPX)
tableCDDOSPX_Close

If intReturn <> 0 Then cmdImport_Monitor

Form_Init


Me.MousePointer = 0
Me.Enabled = True

End Sub


Public Sub cmdContext_Return()
If currentAction = "" Then
    cmdSelect_Click
Else
    SendKeys "{TAB}"
End If

End Sub



Public Sub fgSelect_Reset()
fgSelect_Sort1 = 1: fgSelect_Sort2 = 1
fgSelect_Sort1_Old = 0
fgSelect_RowDisplay = 0: fgSelect_RowClick = 0
fgSelect_arrIndex = 13
blnfgSelect_DisplayLine = False

End Sub

Private Sub txtPageDestinataire_GotFocus()
txt_GotFocus txtPageDestinataire

End Sub


Private Sub txtPageDestinataire_LostFocus()
txt_LostFocus txtPageDestinataire

End Sub


Private Sub txtPageNb_GotFocus()
txt_GotFocus txtPageNb

End Sub

Private Sub txtPageNb_KeyPress(KeyAscii As Integer)
num_KeyAscii KeyAscii

End Sub


Private Sub txtPageNb_LostFocus()
txt_LostFocus txtPageNb

End Sub


Private Sub txtPageNo_GotFocus()
txt_LostFocus txtPageNo

End Sub


Private Sub txtPageNo_KeyPress(KeyAscii As Integer)
num_KeyAscii KeyAscii

End Sub


Private Sub txtPageNo_LostFocus()
txt_LostFocus txtPageNo

End Sub


Private Sub txtPageRéférence_GotFocus()
txt_GotFocus txtPageRéférence

End Sub


Private Sub txtPageRéférence_LostFocus()
txt_LostFocus txtPageRéférence

End Sub


Private Sub txtPays_GotFocus()
txt_GotFocus txtPays

End Sub


Private Sub txtPays_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtPays_LostFocus()
txt_LostFocus txtPays

End Sub



Public Sub cmdImport_Monitor()
Call lstErr_Clear(lstErr, cmdPrint, "CDStat ;Chargement des dossiers"): DoEvents
V = dbCDDOSPX_Import(paramCDSTAT_CDDOSPC, lngNb)
If IsNull(V) Then
    V = "CDStat : " & lngNb & "dossiers importés"
Else
    Call MsgBox(V, vbCritical, "frmCDStat.Msg_Rcv")
    blnImport = False
End If
Call lstErr_AddItem(lstErr, cmdPrint, V): DoEvents

Call lstErr_AddItem(lstErr, cmdPrint, "CDStat ;Chargement des Mouvements"): DoEvents
V = dbCDPOSPX_Import(paramCDSTAT_CDPOSPC, lngNb)
If IsNull(V) Then
    V = "CDStat : " & lngNb & " mvts importés"
Else
    Call MsgBox(V, vbCritical, "frmCDStat.Msg_Rcv")
    blnImport = False
End If
Call lstErr_AddItem(lstErr, cmdPrint, V): DoEvents

End Sub

Public Sub fraSelect_Enabled()

chkCompte.Value = "0": chkCompte.Enabled = False
chkDossier.Value = "0": chkDossier.Enabled = False: txtDossierMax = ""
chkDevise.Value = "0": chkDevise.Enabled = False
chkPays.Value = "0": chkPays.Enabled = False
chkCDE.Value = "1": chkCDE.Enabled = False
chkCDI.Value = "0": chkCDI.Enabled = False
chkAmj.Value = "0": chkAmj.Enabled = False
chkCVAmj = "0": chkCVAmj.Enabled = False: fraCVAmj.Enabled = False
chkConfirmé = "1": chkConfirmé.Enabled = False
chkNotifié = "1": chkNotifié.Enabled = False
chkCVEurVeille = "1": chkCVEurVeille.Enabled = False
chkCVEurHistorique = "0": chkCVEurHistorique.Enabled = False
chkScale = "1": chkScale.Enabled = False
chkGraphique = "1": chkGraphique.Enabled = False
chkBenef.Value = "0": chkBenef.Enabled = False
chkPagePersonnalisée.Value = "0": chkPagePersonnalisée.Enabled = False

fraPrint.Enabled = False
fraSelect.Enabled = True
End Sub

Private Sub txtScaleMax_GotFocus()
txt_GotFocus txtScaleMax

End Sub


Private Sub txtScaleMax_KeyPress(KeyAscii As Integer)
num_KeyAscii KeyAscii
End Sub


Private Sub txtScaleMax_LostFocus()
txt_LostFocus txtScaleMax

End Sub


Private Sub txtScaleStep_GotFocus()
txt_GotFocus txtScaleStep

End Sub

Private Sub txtScaleStep_KeyPress(KeyAscii As Integer)
num_KeyAscii KeyAscii

End Sub


Private Sub txtScaleStep_LostFocus()
txt_LostFocus txtScaleStep

End Sub


