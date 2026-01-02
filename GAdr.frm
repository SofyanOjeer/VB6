VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmGAdr 
   AutoRedraw      =   -1  'True
   Caption         =   "Gestion des adresses"
   ClientHeight    =   6765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9420
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6765
   ScaleWidth      =   9420
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00C0C0FF&
      Caption         =   "en &Attente"
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
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Ok"
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
      TabIndex        =   1
      Top             =   0
      Width           =   1200
   End
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   450
      Left            =   6120
      TabIndex        =   4
      Top             =   0
      Width           =   2745
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   500
      Left            =   8880
      Picture         =   "GAdr.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   6135
      Left            =   0
      TabIndex        =   5
      Top             =   480
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   10821
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   706
      TabCaption(0)   =   "Liste des adresses"
      TabPicture(0)   =   "GAdr.frx":0102
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraSelect"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Mise à jour"
      TabPicture(1)   =   "GAdr.frx":011E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraAdr"
      Tab(1).ControlCount=   1
      Begin VB.Frame fraAdr 
         Height          =   5295
         Left            =   -74880
         TabIndex        =   10
         Top             =   720
         Width           =   9135
         Begin VB.OptionButton optRoutageInterne 
            Caption         =   "expédier le courier"
            Height          =   255
            Left            =   1560
            TabIndex        =   17
            Top             =   1560
            Width           =   2655
         End
         Begin VB.OptionButton optRoutageExterne 
            Caption         =   "retenir le courier"
            Height          =   195
            Left            =   1560
            TabIndex        =   16
            Top             =   1200
            Width           =   2415
         End
         Begin VB.TextBox txtL1 
            Height          =   285
            Left            =   1920
            TabIndex        =   15
            Top             =   840
            Width           =   6975
         End
         Begin VB.TextBox txtL0 
            Height          =   285
            Left            =   1920
            TabIndex        =   14
            Top             =   360
            Width           =   6975
         End
         Begin VB.CheckBox chkL0 
            Alignment       =   1  'Right Justify
            Caption         =   "Nom"
            Height          =   255
            Left            =   240
            TabIndex        =   13
            Top             =   360
            Width           =   1575
         End
         Begin VB.Frame fraAdrD 
            Height          =   2895
            Left            =   120
            TabIndex        =   11
            Top             =   2160
            Width           =   8895
            Begin VB.TextBox txtCP 
               Height          =   285
               Left            =   1800
               TabIndex        =   27
               Top             =   1560
               Width           =   495
            End
            Begin VB.TextBox txtPays 
               Height          =   285
               Left            =   1800
               TabIndex        =   25
               Top             =   2040
               Width           =   495
            End
            Begin VB.TextBox txtL4 
               Height          =   285
               Left            =   3000
               TabIndex        =   24
               Top             =   1560
               Width           =   5775
            End
            Begin VB.TextBox txtL3 
               Height          =   285
               Left            =   1800
               TabIndex        =   23
               Top             =   960
               Width           =   6975
            End
            Begin VB.TextBox txtL2 
               Height          =   285
               Left            =   1800
               TabIndex        =   22
               Top             =   360
               Width           =   6975
            End
            Begin VB.Label libPays 
               Caption         =   "-"
               Height          =   255
               Left            =   3000
               TabIndex        =   26
               Top             =   2160
               Width           =   5655
            End
            Begin VB.Label lblPays 
               Caption         =   "Pays"
               Height          =   255
               Left            =   120
               TabIndex        =   21
               Top             =   2040
               Width           =   975
            End
            Begin VB.Label lblL4 
               Caption         =   "Code postal commune"
               Height          =   495
               Left            =   120
               TabIndex        =   20
               Top             =   1440
               Width           =   1215
            End
            Begin VB.Label lblL3 
               Caption         =   "N° voie"
               Height          =   255
               Left            =   120
               TabIndex        =   19
               Top             =   960
               Width           =   1335
            End
            Begin VB.Label lblL2 
               Caption         =   "Lieu-dit,résidence"
               Height          =   255
               Left            =   120
               TabIndex        =   18
               Top             =   360
               Width           =   1455
            End
         End
         Begin VB.Label lblL1 
            Caption         =   "Complément"
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   840
            Width           =   1575
         End
      End
      Begin VB.Frame fraSelect 
         Height          =   5535
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   9135
         Begin MSFlexGridLib.MSFlexGrid fgSelect 
            Height          =   2610
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   8835
            _ExtentX        =   15584
            _ExtentY        =   4604
            _Version        =   393216
            Rows            =   1
            Cols            =   3
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
            AllowUserResizing=   3
            FormatString    =   "<Référence      |< Nature   |< lien adresse                         "
         End
         Begin MSFlexGridLib.MSFlexGrid fgAdresse 
            Height          =   2250
            Left            =   120
            TabIndex        =   9
            Top             =   3120
            Width           =   8835
            _ExtentX        =   15584
            _ExtentY        =   3969
            _Version        =   393216
            Rows            =   1
            Cols            =   7
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
            AllowUserResizing=   3
            FormatString    =   $"GAdr.frx":013A
         End
      End
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
      Height          =   405
      Left            =   3600
      TabIndex        =   6
      Top             =   0
      Width           =   2535
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuContext 
      Caption         =   "mnuContext"
      Visible         =   0   'False
      Begin VB.Menu mnuAbandonner 
         Caption         =   "Abandonner"
      End
      Begin VB.Menu mnuQuitter 
         Caption         =   "Quitter"
      End
   End
   Begin VB.Menu mnuGAdresse 
      Caption         =   "mnuGAdresse"
      Visible         =   0   'False
      Begin VB.Menu mnuGAdresseAddNew 
         Caption         =   "Créer une adresse"
      End
      Begin VB.Menu mnuGAdresseUpdate 
         Caption         =   "modifier une adresse"
      End
      Begin VB.Menu mnuGAdresseDelete 
         Caption         =   "Effacer une adresse"
      End
   End
   Begin VB.Menu mnuGEntité 
      Caption         =   "mnuGEntité"
      Visible         =   0   'False
      Begin VB.Menu mnuGEntitéAddNew 
         Caption         =   "Créer un lien vers une adresse"
      End
      Begin VB.Menu mnuGEntitéDelete 
         Caption         =   "Effacer un lien vers une adresse"
      End
      Begin VB.Menu mnuGEntitéUpdate 
         Caption         =   "Modifierun lien vers une adresse"
      End
      Begin VB.Menu mnuGEntitéTitulaire 
         Caption         =   "Ajouter un titulaire"
      End
   End
End
Attribute VB_Name = "frmGAdr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim arrTag() As Boolean, arrTagNb As Integer
Dim lastActiveControl_Name  As String, currentActiveControl_Name As String, currentAction As String
Dim MouseMoveActiveControl_Name  As String, MouseMoveActiveControl As typeUsrColor
Dim blnMsgBox_Quit As Boolean, blnAddNew As Boolean, blnGlobalControl As Boolean, blnControl As Boolean
Dim X As String, I As Integer, Msg As String, valX As String, X1 As String, V As Variant, curX As Currency, dblX As Double
Dim GAdresseAut As typeAuthorization

Dim recTable As typeElpTable

Dim fGAdresse_FormatString As String, fGAdresse_K As Integer
Dim fGAdresse_RowDisplay As Integer, fGAdresse_RowClick As Integer
Dim fGAdresse_ColorClick As Long, fGAdresse_ColorDisplay As Long
Dim fGAdresse_Sort1 As Integer, fGAdresse_Sort2 As Integer
Dim fGAdresse_SortAD As Integer, fGAdresse_Sort1_Old As Integer

Dim fgSelect_FormatString As String, fgSelect_K As Integer
Dim fgSelect_RowDisplay As Integer, fgSelect_RowClick As Integer
Dim fgSelect_ColorClick As Long, fgSelect_ColorDisplay As Long
Dim fgSelect_Sort1 As Integer, fgSelect_Sort2 As Integer
Dim fgSelect_SortAD As Integer, fgSelect_Sort1_Old As Integer

Dim recGEntité As typeGEntité, xGEntité As typeGEntité, mGEntité As typeGEntité, mEchéancierGEntité As typeGEntité
Dim arrGAdresse() As typeGAdresse, recGAdresse As typeGAdresse, mGAdresse As typeGAdresse
Dim arrGAdresse_NB As Integer, arrGAdresse_Index As Integer, arrGAdresse_NBMax As Integer
Dim recCompte As typeCompte
Dim blncmdOk_Visible As Boolean, blnErr As Boolean, blncmdSave_Visible As Boolean
Dim blnfgSelect_DisplayLine As Boolean, blnfgEchéance_DisplayLine As Boolean

Public Sub fgSelect_Load()
Dim X As String, mMethod As String

recGEntité_Init xGEntité

Select Case currentAction
    Case "mnuListàValider"
            xGEntité.Method = "SnapLS"
            xGEntité.Application = paramGAdresse_Service
            xGEntité.IdRéférence = "0000000000"
            xGEntité.Statut = "à"
            
            arrGEntité(0) = xGEntité
            arrGEntité(0).IdRéférence = "999999999"

    Case "mnuListGAdresse"
            xGEntité.Method = "SnapLRI"
            X = Trim(txtSelect)
            xGEntité.RéférenceInterne = X
            xGEntité.Application = paramGAdresse_Service
            xGEntité.IdRéférence = 0
            xGEntité.Statut = " "
            
            arrGEntité(0) = xGEntité
            arrGEntité(0).IdRéférence = 999999999
            arrGEntité(0).RéférenceInterne = X & "9z"

End Select

mMethod = Trim(xGEntité.Method)
arrGEntité_NBMax = 0
arrGEntité_Suite = True: arrGEntité_NB = 0
Do Until Not arrGEntité_Suite
    srvGEntité_Monitor xGEntité
    xGEntité = arrGEntité(arrGEntité_NB)
    xGEntité.Method = mMethod & "+"
Loop
fgSelect_Display
End Sub
Private Sub fgSelect_Display()
Dim K2 As Integer, I As Integer
Dim curDB As Currency, curCR As Currency, curX As Currency

SSTab1.Tab = 0

fgSelect.Visible = True
fgSelect.Clear: fgSelect.Rows = 1: fgSelect_RowDisplay = 0

fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
fgSelect.Enabled = True
For arrGEntité_Index = 1 To arrGEntité_NB
    If arrGEntité(arrGEntité_Index).Method <> constIgnore And arrGEntité(arrGEntité_Index).Method <> constDelete Then
        fgSelect.Rows = fgSelect.Rows + 1
        fgSelect.Row = fgSelect.Rows - 1
        fgSelect_DisplayLine
    End If
Next arrGEntité_Index

fgSelect_SortAD = 5
If fgSelect.Rows = 1 Then Exit Sub
'fgSelect_Sort

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


Public Sub fGAdresse_Sort()
If fgAdresse.Rows > 1 Then
    fgAdresse.Row = 1
    fgAdresse.RowSel = fgAdresse.Rows - 1
    If fGAdresse_Sort1_Old = fGAdresse_Sort1 Then
        If fGAdresse_SortAD = 5 Then
            fGAdresse_SortAD = 6
        Else
            fGAdresse_SortAD = 5
        End If
    Else
        fGAdresse_SortAD = 5
    End If
    fGAdresse_Sort1_Old = fGAdresse_Sort1
    
    fgAdresse.Col = fGAdresse_Sort1
    fgAdresse.ColSel = fGAdresse_Sort2
    fgAdresse.Sort = fGAdresse_SortAD
End If
    

End Sub


'---------------------------------------------------------
Public Sub arrTag_Set(ByVal B As Boolean)
'---------------------------------------------------------
For I = 0 To arrTagNb: arrTag(I) = B: Next I
End Sub


Public Sub cmdContext_Quit()
blnControl = False
    If currentAction <> "" Then
        currentAction = ""
        cmdContext.Caption = constcmdRechercher
        fgSelect.Enabled = True
        fgAdresse.Enabled = True
        cmdReset
        SSTab1.Tab = 0
    Else
        If blnMsgBox_Quit Then
        X = MsgBox("Voulez-vous réellement abandonner?", vbYesNo + vbQuestion + vbDefaultButton2, "Saisie non enregistrée")
        Else
           X = vbYes
        End If
        If X = vbYes Then Unload Me
    End If

End Sub
Public Sub cmdContext_Return()
If fraOption.Visible Then
    fraOption.Visible = False
Else
    If SSTab1.Tab = 0 And Trim(txtSelect) <> "" Then
        mnuListGAdresse_Click
    Else
        SendKeys "{TAB}"
    End If
End If

End Sub


'---------------------------------------------------------
Private Sub cmdQuit_Click()
'---------------------------------------------------------
Unload Me

End Sub

'---------------------------------------------------------
Public Sub cmdReset()
'---------------------------------------------------------
blnControl = False
usrColor_Set
cmdOk.Caption = constàValider: cmdOk.Visible = False
cmdSave.Caption = constEnAttente: cmdSave.Visible = False
cmdContext.Caption = constcmdRechercher: blnMsgBox_Quit = False
arrTag_Set False
libRéférenceInterne = ""
lstErr.Visible = False
blnComptaAuto = False
cmdOk.FontSize = 8: cmdOk.FontName = "MS Sans Serif"
blncmdOk_Visible = False: blncmdSave_Visible = False
blnfgSelect_DisplayLine = False: blnfgEchéance_DisplayLine = False

fraGAdresse.Enabled = False
fgAdresse.Clear: fgAdresse.Rows = 1: fGAdresse_RowDisplay = 0
If cboNature.ListCount > -1 Then cboNature.ListIndex = 0
CV1 = CV_Euro
CV1.DeviseIso = "FRF"
CV_Attribut CV1
txtDevise = CV1.DeviseIso
txtCapital = ""
txtTaux = ""
txtCommissionFlat = ""
optTrimestriel = True
chkCommissionPériodique = "1"
fraCommissionPériodique.Enabled = True
optComPériodiqueTaux = True

lblAMJEffet.Visible = False: txtAMJEffet.Visible = False
Call lbl_Style(lblCapital, False)
Call lbl_Style(lblAMJEffet, False)
Call lbl_Style(lblAmjFin, False)
Call lbl_Style(lblPréavisNbj, False)
Call chk_Style(chkMainLevée, False)


mAMJReprise = DSys
wAmjEngagement = DSys: Call DTPicker_Set(txtAmjEngagement, wAmjEngagement)
wAmjFin = DSys: Call DTPicker_Set(txtAmjEngagement, wAmjFin)
wAmjEchéance = dateFinDeMois(dateElp("MoisAdd", 1, DSys)): Call DTPicker_Set(txtAmjEchéance, wAmjEchéance)
txtDonneurDordre = ""
recRacineInit C_Racine
mDonneurDordre = "": mCboNature = ""
txtRéférenceInterne = ""
txtRéférenceExterne = ""
txtEngagementCompte = "": libEngagementCompte = ""
txtEchéanceCompte = "": libEchéanceCompte = ""
optEchéanceAnniversaire = True
recGEntité_Init mGEntité
mGEntité.Statut = "à"
mGEntité.StatutPlus = "?"
mGEntité.Method = constAddNew
mEChéanceCompte = Space$(11): mEngagementCompte = Space$(11): mEngagementCorrCompte = Space$(11)
Call DTPicker_Set(txtAmjEngagement, DSys)
Call DTPicker_Set(txtAMJFin, DSys)
chkComptaReprise = "0"
chkMainLevée = "0"
chkComptaReprise = "0"

fraOption.Visible = False
blnEchéancier_Gen = False
recGEntité_Init mEchéancierGEntité: mEchéancierGEntité.Application = paramGAdresse_Service
saveGAdresse_Index_GA02 = 0
blnControl = True
End Sub



Public Sub Form_Init()

GAdresse_Compta.Param_Init mNature, cboNature
libRéférenceInterne.ForeColor = vbBlue

recElpTable_Init recElpTable
recElpTable.Method = "Seek="
recElpTable.Id = "GEntité" '"Param"
recElpTable.K1 = mNature
recElpTable.K2 = "ComàRéclamer"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramGAdresse_CompteCommissionàRéclamer = mId$(recElpTable.Memo, 1, 11)
If Not IsNumeric(paramGAdresse_Service) Then GoTo Num_Error

wAmjEchéanceTrt = dateElp("Jour", 15, DSys)
Call DTPicker_Set(txtAmjMax, wAmjEchéanceTrt)
SSTab1.Tab = 0
tableElpTable_Open
paramAmjEngagementMin = paramAmjOpérationMin   ' jpl 2000-09-01 mId$(DSys, 1, 6) & "01"
paramAmjEngagementMax = dateElp("Ouvré", 7, DSys)
ReDim arrGEntité(1)
cmdReset
mnuGAdresseSaisir.Enabled = GAdresseAut.Saisir
mnuListàValider.Enabled = GAdresseAut.Consulter
mnuListGAdresse.Enabled = GAdresseAut.Consulter
mnuComptaEchéancier.Enabled = GAdresseAut.Consulter
mnuListEchéancier.Enabled = GAdresseAut.Consulter
mnuComptaLotsàValider.Enabled = GAdresseAut.Comptabiliser
mnuLotComptabilisé_Annuler.Enabled = GAdresseAut.Xspécial
blnControl = False
''txtSelect.SetFocus
fgSelect_Sort1 = 0: fgSelect_Sort2 = 0
fgSelect_Sort1_Old = 0
fgSelect_FormatString = fgSelect.FormatString
fgSelect_RowDisplay = 0: fgSelect_RowClick = 0

fGAdresse_Sort1 = 11: fGAdresse_Sort2 = 11
fGAdresse_Sort1_Old = 11
fGAdresse_FormatString = fgAdresse.FormatString
fGAdresse_RowDisplay = 0: fGAdresse_RowClick = 0

Exit Sub

Table_Error:
MsgBox recElpTable.Id & " : " & recElpTable.K1 & " : " & recElpTable.K2 & " : Table", vbCritical, "frmGAdresse.Form_Init"
Exit Sub

Memo_Error:
MsgBox recElpTable.Id & " : " & recElpTable.K1 & " : " & recElpTable.K2 & " : Mémo absent", vbCritical, "frmGAdresse.Form_Init"
Exit Sub

Num_Error:
MsgBox recElpTable.Id & " : " & recElpTable.K1 & " : " & recElpTable.K2 & " : " & recElpTable.Memo & " :Mémo non numérique", vbCritical, "GAdresseEspèces_Param_Init"
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

Public Sub Msg_Rcv(Msg As String)
'---------------------------------------------------------
End Sub

Public Sub Msg_Snd(ByVal X As String)
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


Private Sub chkL0_Click()
If blnControl Then cmdControl

End Sub

Private Sub chkL0_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkL0

End Sub


Private Sub cmdContext_Click()
Select Case cmdContext.Caption
    Case Is = constcmdRechercher: Me.PopupMenu mnuContext, vbPopupMenuLeftButton
    Case Is = constcmdAbandonner: cmdContext_Quit
End Select

End Sub


Private Sub cmdContext_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdContext
End Sub

Private Sub cmdOk_Click()
Dim blnPrint As Boolean, wPrint_Msg As String
Dim V

blnPrint = False
wPrint_Msg = "GAdresse"

If cmdOk.Caption = constàCompta Then
    cmdSave_àCompta
Else
    cmdControl
    If lstErr.ListCount <> 0 Then Exit Sub
    frmGAdresse.Enabled = False
    Select Case cmdOk.Caption
        Case constàValider
'            recGEntité.Statut = "à"
'            recGEntité.StatutPlus = "V "
'            recGEntité.MajAMJ = DSys
'            recGEntité.MajHMS = time_Hms
'            recGEntité.MajUsr = usrId
'            wPrint_Msg = constàValider 'cmdPrint_Call constàValider
'            blnPrint = True
    Case Else
            Call lstErr_AddItem(lstErr, cmdContext, "? cmdOk : " & cmdOk.Caption)
    End Select

'    If lstErr.ListCount = 0 Then
    End If
    
    frmGAdresse.Enabled = True
    AppActivate frmGAdresse.Caption
End If
End Sub

Private Sub cmdPrint_Click()
'Me.PopupMenu mnucmdPrint, vbPopupMenuLeftButton
End Sub

Private Sub cmdSave_Click()
cmdControl
lstErr.Clear
frmGAdresse.Enabled = False
Select Case cmdSave.Caption
        Call lstErr_AddItem(lstErr, cmdContext, "? cmdsave : " & cmdSave.Caption)
End Select

If lstErr.ListCount = 0 Then cmdSave_Db
frmGAdresse.Enabled = True
End Sub
Private Sub fGAdresse_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Y <= fgAdresse.RowHeightMin Then
    Select Case fgAdresse.Col
        Case 0: fGAdresse_Sort1 = 0: fGAdresse_Sort2 = 0: fGAdresse_Sort
        Case 1: fGAdresse_Sort1 = 1: fGAdresse_Sort2 = 1: fGAdresse_Sort
        Case 2: fGAdresse_Sort1 = 2: fGAdresse_Sort2 = 2: fGAdresse_Sort
        Case 3: fGAdresse_Sort1 = 3: fGAdresse_Sort2 = 3: fGAdresse_Sort
        Case 4: fGAdresse_Sort1 = 4: fGAdresse_Sort2 = 4: fGAdresse_Sort
        Case 5: fGAdresse_Sort1 = 5: fGAdresse_Sort2 = 5: fGAdresse_Sort
        Case 6: fGAdresse_Sort1 = 6: fGAdresse_Sort2 = 6: fGAdresse_Sort
    End Select
Else
    fGAdresse_K = fgAdresse.Row * fgAdresse.Cols
    If fgAdresse.Rows > 1 Then
        Call fGAdresse_Color(fGAdresse_RowClick, MouseMoveUsr.BackColor, fGAdresse_ColorClick)
        arrGAdresse_Index = Val(fgAdresse.TextArray(fgAdresse.Cols - 1 + fGAdresse_K))
        '''recGAdresse.CptMvtLot = arrGAdresse(arrGAdresse_Index).CptMvtLot
        recGAdresse = arrGAdresse(arrGAdresse_Index)
        Param_CodeOpération recGAdresse.CodeOpération
        
        If currentAction = constDisplay Then
            mnuEchéancier_Set
            Me.PopupMenu mnuEchéancier, vbPopupMenuLeftButton
           Else
            If recGAdresse.CptMvtLot > 0 Then
                mnuLotàComptaValidation = False
                mnuLotàComptaAnnulation = False
                mnuLotàComptaAnnulation = False
              
                If recGAdresse.Statut = "à" And recGAdresse.StatutPlus = "C " Then
                    mnuLotàComptaValidation = GAdresseAut.Comptabiliser
                    mnuLotàComptaAnnulation = GAdresseAut.Comptabiliser
                    mnuLotàComptaPrint = GAdresseAut.Comptabiliser
                End If
        
                Me.PopupMenu mnuLot, vbPopupMenuLeftButton
            End If
        End If
    End If
End If

End Sub


Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim xStatut As String
If Y <= fgSelect.RowHeightMin Then
    Select Case fgSelect.Col
        Case 0: fgSelect_Sort1 = 0: fgSelect_Sort2 = 0: fgSelect_Sort
        Case 1:  fgSelect_SortX 1
        Case 2: fgSelect_SortX 2
        Case 3, 13: fgSelect_Sort1 = 13: fgSelect_Sort2 = 13: fgSelect_Sort
        Case 4, 14: fgSelect_Sort1 = 14: fgSelect_Sort2 = 14: fgSelect_Sort
        Case 5: fgSelect_Sort1 = 5: fgSelect_Sort2 = 5: fgSelect_Sort
        Case 6: fgSelect_Sort1 = 6: fgSelect_Sort2 = 6: fgSelect_Sort
        Case 7: fgSelect_Sort1 = 7: fgSelect_Sort2 = 7: fgSelect_Sort
        Case 8: fgSelect_Sort1 = 8: fgSelect_Sort2 = 8: fgSelect_Sort
        Case 9: fgSelect_Sort1 = 9: fgSelect_Sort2 = 9: fgSelect_Sort
        Case 10: fgSelect_Sort1 = 10: fgSelect_Sort2 = 12: fgSelect_Sort
        Case 11: fgSelect_Sort1 = 11: fgSelect_Sort2 = 12: fgSelect_Sort
        Case 12: fgSelect_Sort1 = 12: fgSelect_Sort2 = 12: fgSelect_Sort
        Case 16:  fgSelect_SortX 16
    End Select
Else

    fgSelect_K = fgSelect.Row * fgSelect.Cols
    If fgSelect.Rows > 1 Then
        Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
        arrGEntité_Index = Val(fgSelect.TextArray(fgSelect.Cols - 1 + fgSelect_K))
        xGEntité = arrGEntité(arrGEntité_Index)
    
        If xGEntité.IdRéférence > 0 Then
            mnuGAdresseDisplay = GAdresseAut.Consulter
            mnuGAdresseModifier = False
            mnuGAdresseAnnuler = False
            mnuGAdresseEffacer = False
            mnuGAdresseValider = False
            mnuGAdresseAMJFin = False
            mnuGAdresseMainLevéePartielle = False
            mnuGAdresseMainLevée = False
          
            xStatut = xGEntité.Statut & xGEntité.StatutPlus
            If xStatut = "à? " Then
                mnuGAdresseModifier = GAdresseAut.Saisir
                mnuGAdresseEffacer = GAdresseAut.Saisir
            End If
            If xStatut = "àV " Then
              If Not GAdresseAut.Xspécial And Trim(recGEntité.MajUsr) = Trim(usrId) Then
                    Call lstErr_Clear(lstErr, cmdContext, "! Vous ne pouvez pas valider vos opérations")
                Else
                    mnuGAdresseValider = GAdresseAut.Valider
                End If
            End If
            If xStatut = "   " Then
                mnuGAdresseAMJFin = GAdresseAut.Saisir
                mnuGAdresseMainLevéePartielle = GAdresseAut.Saisir
                mnuGAdresseMainLevée = GAdresseAut.Saisir
           End If
    
            Me.PopupMenu mnuGAdresse, vbPopupMenuLeftButton
        End If
    End If
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
Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)
blnControl = False
fgSelect.Clear: fgSelect.Row = 0
fgAdresse.Clear: fgAdresse.Row = 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset
End Sub

Private Sub fraAdr_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub fraAdrD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub fraSelect_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub mnuAbandonner_Click()
cmdContext_Quit
End Sub



Private Sub mnuQuitter_Click()
Unload Me
End Sub







Public Sub cmdControl()
Dim X As String, wMensualité As Currency, wAmj As String, wTaux As Double
Dim wTA As Double, wTEG As Double

If Not Me.Enabled Then Exit Sub
Me.Enabled = False

cmdOk.Visible = False
cmdSave.Visible = False
blnControl = False

lstErr.Clear
lstErr.Height = 200

recGEntité = mGEntité
Select Case currentAction
    Case constValider
            V = fctGEntité_Compare(recGEntité, mGEntité)
            If Not IsNull(V) Then
                Call MsgBox("L'enregistrement après contrôle est différent de l'enregistrement lu :" & Chr$(13) & V, vbCritical, "me : cmdControl")
                Call lstErr_AddItem(lstErr, cmdContext, "? Erreur Contrôle validation")
            End If
End Select

If lstErr.ListCount = 0 Then
    cmdOk.Visible = blncmdOk_Visible
Else
'    SSTab1.Tab = 2
End If

ExitSub:

Me.Enabled = True
If cmdOk.Visible Then cmdOk.SetFocus
    
blnControl = True


End Sub

Public Sub fGAdresse_Display(Fct As String)
Dim I As Integer

fgAdresse.Visible = True
fgAdresse.Clear: fGAdresse_RowDisplay = 0: fGAdresse_RowClick = 0
totalCapital = 0: totalIntérêts = 0

fgAdresse.Rows = 1
fgAdresse.FormatString = fGAdresse_FormatString
fgAdresse.Enabled = True
For arrGAdresse_Index = 1 To arrGAdresse_NB
    recGAdresse = arrGAdresse(arrGAdresse_Index)
    fgAdresse.Rows = fgAdresse.Rows + 1
    fgAdresse.Row = fgAdresse.Rows - 1
    fGAdresse_DisplayLine
Next arrGAdresse_Index

fGAdresse_K = fgAdresse.Cols
 
fGAdresse_SortAD = 5
If fgAdresse.Rows > 1 Then fGAdresse_SortX 9
End Sub

Public Sub fGAdresse_DisplayLine()
Dim K2 As Integer

fGAdresse_K = (fgAdresse.Row) * fgAdresse.Cols
If mEchéancierGEntité.IdRéférence <> recGAdresse.IdRéférence Then
    mEchéancierGEntité.RéférenceInterne = ""
    mEchéancierGEntité.IdRéférence = recGAdresse.IdRéférence: srvGEntité_Find mEchéancierGEntité
End If

fgAdresse.TextArray(0 + fGAdresse_K) = mEchéancierGEntité.RéférenceInterne
fgAdresse.TextArray(1 + fGAdresse_K) = GAdresse_Compta.Param_CodeOpération(recGAdresse.CodeOpération)
fgAdresse.TextArray(2 + fGAdresse_K) = Format(recGAdresse.Capital + recGAdresse.Intérêts, "#### ### ###.00 ")
fgAdresse.TextArray(3 + fGAdresse_K) = dateImp(recGAdresse.AmjEchéanceTrt)
fgAdresse.TextArray(4 + fGAdresse_K) = recStatut_Libellé(recGAdresse.Statut & recGAdresse.StatutPlus)
fgAdresse.TextArray(5 + fGAdresse_K) = Format(recGAdresse.Taux, "#0.00000 ") & recGAdresse.TauxProvisoire
fgAdresse.TextArray(6 + fGAdresse_K) = "du " & dateImp(recGAdresse.AmjDébut) & " au " & dateImp(recGAdresse.AmjFin) & "   (" & recGAdresse.Nbj & "j)"
fgAdresse.TextArray(7 + fGAdresse_K) = dateImp(recGAdresse.CptMvtAMJ) & " " & timeImp(recGAdresse.CptMvtHMS) & " " & recGAdresse.CptMvtUsr
If recGAdresse.CptMvtPièce <> 0 Then
    fgAdresse.TextArray(8 + fGAdresse_K) = "Pièce : " & Format(recGAdresse.CptMvtPièce, "### ### ") & "." & Format(recGAdresse.CptMvtLigne, "### ### ") & " Lot : " & Format(recGAdresse.CptMvtLot, "### ### ")
End If
fgAdresse.TextArray(9 + fGAdresse_K) = Format(recGAdresse.IdRéférence, "### ##0 ") & "_" & Format(recGAdresse.IdSéquence, "### ### ")
fgAdresse.TextArray(10 + fGAdresse_K) = recGAdresse.AmjEchéanceTrt
fgAdresse.TextArray(fgAdresse.Cols - 1 + fGAdresse_K) = arrGAdresse_Index

If recGAdresse.CodeOpération <> "GA01" Then
    fgAdresse.Col = 2: fgAdresse.CellForeColor = errUsr.ForeColor
End If
If recGAdresse.Statut = "A" Then fgAdresse.Col = 4: fgAdresse.CellForeColor = errUsr.ForeColor


End Sub

Public Function Compte_Load(mCompteNuméro As String)
Compte_Load = Null
recCompteInit recCompte
recCompte.Société = SocId$
recCompte.Agence = SocAgence$
recCompte.Devise = CV1.DeviseN
recCompte.Numéro = mCompteNuméro
recCompte.BiaTyp = "000"
recCompte.BiaNum = "00"
recCompte.Method = "SeekL1"
If Not IsNull(srvCompteFind(recCompte)) Then Call lstErr_AddItem(lstErr, lstErr, "? compte inconnu : " & mCompteNuméro): Compte_Load = "?": Exit Function

If recCompte.Situation <> " " Then
    Select Case recCompte.Situation
        Case "B": Call lstErr_AddItem(lstErr, lstErr, " ? Compte bloqué : " & mCompteNuméro): Compte_Load = "?"
        Case "A": Call lstErr_AddItem(lstErr, lstErr, " ? Compte annulé : " & mCompteNuméro): Compte_Load = "?"
        Case Else: Call lstErr_AddItem(lstErr, lstErr, " ? Situation du compte : " & mCompteNuméro): Compte_Load = "?"
    End Select
End If

End Function

Public Sub fraGAdresse_Load(Fct As String)
'2000-01-04 cmdReset
fgSelect_RowClick = 0
Call fgSelect_Color(fgSelect_RowDisplay, vbCyan, fgSelect_ColorClick) 'txtUsr.BackColor)
blnControl = False
xGEntité.Method = "SeekP0"
V = srvGEntité_Monitor(xGEntité)
If IsNull(V) Then
    libRéférenceInterne = Trim(xGEntité.RéférenceInterne) & "_" & Compte_Imp(xGEntité.EngagementCompte)
    blnAmjEchéance = True
    SSTab1.Tab = 1
    mGEntité = xGEntité
    mGEntité.Method = Fct
    mCboNature = mGEntité.Nature
    cbo_Scan mGEntité.Nature, cboNature
    mDonneurDordre = mId$(mGEntité.EngagementCompte, 1, 5)
    txtDonneurDordre = mDonneurDordre
    txtEngagementCompte = Compte_Display(mGEntité.EngagementCompte)
    txtDevise = mGEntité.Devise
    txtCapital = Format$(mGEntité.Capital, "### ### ### ##0.00")
    txtRéférenceInterne = Trim(mGEntité.RéférenceInterne)
    txtRéférenceExterne = Trim(mGEntité.RéférenceExterne)
    Call DTPicker_Set(txtAmjEngagement, mGEntité.AmjDébut): wAmjEngagement = mGEntité.AmjDébut
    Call DTPicker_Set(txtAMJFin, mGEntité.AmjFin): wAmjFin = mGEntité.AmjFin
    
    If mGEntité.PréavisNbj = 999 Then
        chkMainLevée = "1"
    Else
        chkMainLevée = "0"
        txtPréavisNbj = mGEntité.PréavisNbj
    End If
    
    If Trim(mGEntité.EchéanceCompte) <> "" Then 'paramGAdresse_CompteCommissionàRéclamer Then
        chkCommissionàRéclamer = "0"
        txtEchéanceCompte = Compte_Display(mGEntité.EchéanceCompte)
    Else
        chkCommissionàRéclamer = "1"
        txtEchéanceCompte = ""
    End If
    
    If mGEntité.AmjEchéance1 = mGEntité.AmjDébut Then
        chkAmjEchéance = "0"
    Else
        chkAmjEchéance = "1"
    End If
    
    Call DTPicker_Set(txtAmjEchéance, mGEntité.AmjEchéance1): wAmjEchéance = mGEntité.AmjEchéance1
    If mGEntité.AmjEchéanceS = "M" Then
        optEchéanceFinDeMois = True
    Else
        optEchéanceAnniversaire = True
    End If
    If mGEntité.TauxMarge <> 0 Then
        chkCommissionPériodique = "1"
        If Trim(mGEntité.TauxRéférence) <> "Montant" Then
            txtTaux = Format$(mGEntité.TauxMarge, "#0.00000")
            optComPériodiqueTaux = True
        Else
            optComPériodiqueMontant = True
            txtTaux = Format$(mGEntité.TauxMarge, "##### ##0.00")
       End If
    Else
        chkCommissionPériodique = "0"
        txtTaux = ""
    End If
    
    Select Case mGEntité.Périodicité
        Case "M": optMensuel = True: fctPériodicité = "MoisAdd"
        Case "T": optTrimestriel = True: fctPériodicité = "TrimestreAdd"
        Case "S": optSemestriel = True: fctPériodicité = "SemestreAdd"
        Case "A": optAnnuel = True: fctPériodicité = "AnAdd"
        Case Else: optMensuel = True: fctPériodicité = "MoisAdd"
   End Select
      
    If mGEntité.Frais = 0 Then
        chkCommissionFlat = "0"
    Else
        chkCommissionFlat = "1"
        txtCommissionFlat = Format$(mGEntité.Frais, "### ### ### ##0.00")
    End If
    
    mAMJReprise = mGEntité.MajAMJ
    If mGEntité.optReprise = "R" Then
        chkComptaReprise = "1"
    Else
        chkComptaReprise = "0"
    End If
    libStatut = "Statut         : " & recStatut_Libellé(mGEntité.Statut & mGEntité.StatutPlus) & Chr$(13) _
                & "Référence  : " & Format$(mGEntité.IdRéférence, "#### ### ##0") & Chr$(13) & Chr$(13) _
                & "Saisi par  : " & mGEntité.MajUsr & Chr$(13) _
                & "                 : " & dateImp(mGEntité.MajAMJ) & " " & timeImp(mGEntité.MajHMS) & Chr$(13) & Chr$(13) _
                & "Validé par :" & mGEntité.ValUsr & Chr$(13) _
                & "                 : " & dateImp(mGEntité.valAMJ) & " " & timeImp(mGEntité.ValHMS)
   cmdControl
    
    If mGEntité.Statut = "à" Then
        Select Case mGEntité.StatutPlus
            Case Is = "V ": cmdOk.Caption = constValider
                            cmdSave.Caption = constàModifier
         
            Case Is = "? "
                        If Fct = constEffacer Then
                            cmdSave.Caption = constEffacer
                        Else
                            cmdOk.Caption = constàValider
                        End If
        End Select
    Else
        fGAdresse_Load
    End If
    cmdSave.Visible = blncmdSave_Visible
    blnfgSelect_DisplayLine = True
End If

End Sub
Public Sub fgSelect_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
mRow = fgSelect.Row

If lRow > 0 Then
    fgSelect.Row = lRow
    For I = 0 To 16
        fgSelect.Col = I: fgSelect.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgSelect.Row = mRow
    If fgSelect.Row > 0 Then
        lRow = fgSelect.Row
        lColor_Old = fgSelect.CellBackColor
        For I = 0 To 16
          fgSelect.Col = I: fgSelect.CellBackColor = lColor
        Next I
        fgSelect.Col = 0
    End If
End If

End Sub

Public Sub fGAdresse_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
mRow = fgAdresse.Row


If lRow > 0 Then
    fgAdresse.Row = lRow
    For I = 0 To fgAdresse.Cols - 1
        fgAdresse.Col = I: fgAdresse.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgAdresse.Row = mRow
    If fgAdresse.Row > 0 Then
        lRow = fgAdresse.Row
        lColor_Old = fgAdresse.CellBackColor
        For I = 0 To fgAdresse.Cols - 1
          fgAdresse.Col = I: fgAdresse.CellBackColor = lColor
        Next I
        fgAdresse.Col = 0
    End If
End If

End Sub

Public Sub fGAdresse_AddNew()
V = srvGAdresse_Dtaq_Put("Init", recGAdresse)
If Not IsNull(V) Then fGAdresse_Delete: Exit Sub
For I = 1 To arrGAdresse_NB
'    arrGAdresse(I).Method = constAddNew
'    arrGAdresse(I).IdRéférence = recGEntité.IdRéférence
    recGAdresse = arrGAdresse(I)
'   If recGEntité.optReprise = "R" And recGAdresse.AmjEchéanceTrt < DSys Then
'        recGAdresse.Statut = "R"
'        recGAdresse.StatutPlus = "ep"
'    Else
'        Param_CodeOpération recGAdresse.CodeOpération
'        If paramGAdresse_CodeOpération_Compta <> "A" Then recGAdresse.Statut = "M": recGAdresse.StatutPlus = "an"
'    End If
'     arrGAdresse(I) = recGAdresse
    V = srvGAdresse_Dtaq_Put("Add", recGAdresse)
    If Not IsNull(V) Then fGAdresse_Delete: Exit Sub
Next I
V = srvGAdresse_Dtaq_Put("Snd", recGAdresse)
If Not IsNull(V) Then fGAdresse_Delete: Exit Sub

End Sub

Public Sub fGAdresse_Update()
V = srvGAdresse_Dtaq_Put("Init", arrGAdresse(1))
If Not IsNull(V) Then: Exit Sub
For I = 1 To arrGAdresse_NB
    If arrGAdresse(I).Method = constUpdate Or arrGAdresse(I).Method = constAddNew Or arrGAdresse(I).Method = constDelete Then
        V = srvGAdresse_Dtaq_Put("Add", arrGAdresse(I))
        If Not IsNull(V) Then Exit Sub
    End If
Next I
V = srvGAdresse_Dtaq_Put("Snd", arrGAdresse(1))
If Not IsNull(V) Then Exit Sub

End Sub

Public Sub fGAdresse_Delete()
Call lstErr_AddItem(lstErr, cmdContext, V)
Call lstErr_AddItem(lstErr, cmdContext, "!! Suppression de l'échéancier")
arrGAdresse(1).Method = "DeleteAll"
Call srvGAdresse_Update(arrGAdresse(1))
End Sub

Public Sub fGAdresse_Load()
ReDim arrGAdresse(1)

recGAdresse_Init recGAdresse
recGAdresse.Method = "SnapP0"
recGAdresse.IdRéférence = mGEntité.IdRéférence

arrGAdresse(0) = recGAdresse
arrGAdresse(0).IdSéquence = 999

Call srvGAdresse_Load(recGAdresse, arrGAdresse(0))
arrGAdresse_NB = srvGAdresse.arrGAdresse_NB
ReDim arrGAdresse(arrGAdresse_NB)
For I = 1 To arrGAdresse_NB
    arrGAdresse(I) = srvGAdresse.arrGAdresse(I)
Next I

fGAdresse_Display "T"

End Sub

Public Sub fgSelect_DisplayLine()
fgSelect_K = (fgSelect.Row) * fgSelect.Cols
fgSelect.TextArray(7 + fgSelect_K) = arrGEntité(arrGEntité_Index).Nature
fgSelect.TextArray(1 + fgSelect_K) = Format(arrGEntité(arrGEntité_Index).Capital, "#### ### ###.00 ")
fgSelect.TextArray(2 + fgSelect_K) = arrGEntité(arrGEntité_Index).Devise
fgSelect.TextArray(3 + fgSelect_K) = dateImp(arrGEntité(arrGEntité_Index).AmjDébut)
fgSelect.TextArray(4 + fgSelect_K) = dateImp(arrGEntité(arrGEntité_Index).AmjFin)
fgSelect.TextArray(5 + fgSelect_K) = Compte_Imp(arrGEntité(arrGEntité_Index).EngagementCompte)
Call CV_AttributS(arrGEntité(arrGEntité_Index).Devise, CV1)
recCompteInit recCompte
recCompte.Société = SocId$
recCompte.Agence = SocAgence$
recCompte.Devise = CV1.DeviseN
recCompte.Numéro = arrGEntité(arrGEntité_Index).EngagementCompte
mdbCptP0_Find recCompte
'Call Compte_Load(arrGEntité(arrGEntité_Index).EngagementCompte)
fgSelect.TextArray(6 + fgSelect_K) = recCompte.Intitulé
fgSelect.TextArray(0 + fgSelect_K) = arrGEntité(arrGEntité_Index).RéférenceInterne
fgSelect.TextArray(8 + fgSelect_K) = arrGEntité(arrGEntité_Index).RéférenceExterne
fgSelect.TextArray(9 + fgSelect_K) = ""
fgSelect.TextArray(9 + fgSelect_K) = recStatut_Libellé(arrGEntité(arrGEntité_Index).Statut & arrGEntité(arrGEntité_Index).StatutPlus)
fgSelect.TextArray(10 + fgSelect_K) = arrGEntité(arrGEntité_Index).MajUsr & " " & dateImp(arrGEntité(arrGEntité_Index).MajAMJ) & " " & timeImp(arrGEntité(arrGEntité_Index).MajHMS)
fgSelect.TextArray(11 + fgSelect_K) = arrGEntité(arrGEntité_Index).ValUsr & " " & dateImp(arrGEntité(arrGEntité_Index).valAMJ) & " " & timeImp(arrGEntité(arrGEntité_Index).ValHMS)
fgSelect.TextArray(12 + fgSelect_K) = arrGEntité(arrGEntité_Index).IdRéférence
fgSelect.TextArray(13 + fgSelect_K) = arrGEntité(arrGEntité_Index).AmjDébut
fgSelect.TextArray(14 + fgSelect_K) = arrGEntité(arrGEntité_Index).AmjFin
fgSelect.TextArray(15 + fgSelect_K) = ""
fgSelect.TextArray(16 + fgSelect_K) = arrGEntité_Index

End Sub

Public Sub fgSelect_SortX(lK As Integer)
Dim I As Integer
For I = 1 To fgSelect.Rows - 1
    fgSelect.Row = I
    fgSelect.Col = fgSelect.Cols - 1
    arrGEntité_Index = Val(fgSelect.Text)
    fgSelect.Col = 15
    Select Case lK
        Case 1: fgSelect.Text = Format$(arrGEntité(arrGEntité_Index).Capital, "000000000000000.00") & arrGEntité(arrGEntité_Index).Devise
        Case 2: fgSelect.Text = arrGEntité(arrGEntité_Index).Devise & Format$(arrGEntité(arrGEntité_Index).Capital, "000000000000000.00")
        Case 16: fgSelect.Text = Format$(arrGEntité_Index, "0000000000")
    End Select
Next I

fgSelect_Sort1 = 15: fgSelect_Sort2 = 15
fgSelect_Sort
End Sub
Public Sub fGAdresse_SortX(lK As Integer)
Dim I As Integer
For I = 1 To fgAdresse.Rows - 1
    fgAdresse.Row = I
    fgAdresse.Col = fgAdresse.Cols - 1
    arrGAdresse_Index = Val(fgAdresse.Text)
    fgAdresse.Col = 10
    Select Case lK
        Case 2: fgAdresse.Text = Format$(arrGAdresse(arrGAdresse_Index).Capital, "000000000000000.00")
        Case 3: fgAdresse.Text = arrGAdresse(arrGAdresse_Index).AmjEchéanceTrt
        Case 9: fgAdresse.Text = Format(recGAdresse.IdRéférence, "000000") & "_" & Format(recGAdresse.IdSéquence, "0000000")
        Case 11: fgAdresse.Text = Format$(arrGAdresse_Index, "0000000000")
    End Select
Next I

fGAdresse_Sort1 = 10: fGAdresse_Sort2 = 10
fGAdresse_Sort
End Sub


Private Sub optRoutageExterne_Click()
If blnControl Then cmdControl

End Sub


Private Sub optRoutageExterne_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optRoutageExterne
End Sub


Private Sub optRoutageInterne_Click()
If blnControl Then cmdControl

End Sub


Private Sub optRoutageInterne_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set optRoutageInterne
End Sub


Private Sub txtCP_GotFocus()
txt_GotFocus txtCP
End Sub

Private Sub txtCP_LostFocus()
txt_LostFocus txtCP
If blnControl Then cmdControl


End Sub


Private Sub txtL0_GotFocus()
txt_GotFocus txtL0

End Sub

Private Sub txtL0_LostFocus()
txt_LostFocus txtL0
If blnControl Then cmdControl

End Sub


Private Sub txtL1_GotFocus()
txt_GotFocus txtL1

End Sub

Private Sub txtL1_LostFocus()
txt_LostFocus txtL1
If blnControl Then cmdControl

End Sub


Private Sub txtL2_GotFocus()
txt_GotFocus txtL2

End Sub

Private Sub txtL2_LostFocus()
txt_LostFocus txtL2
If blnControl Then cmdControl


End Sub

Private Sub txtL3_GotFocus()
txt_GotFocus txtL3

End Sub

Private Sub txtL3_LostFocus()
txt_LostFocus txtL3
If blnControl Then cmdControl


End Sub

Private Sub txtL4_GotFocus()
txt_GotFocus txtL4

End Sub

Private Sub txtL4_LostFocus()
txt_LostFocus txtL4
If blnControl Then cmdControl


End Sub

Private Sub txtPays_GotFocus()
txt_GotFocus txtPays

End Sub

Private Sub txtPays_LostFocus()
txt_LostFocus txtPays
If blnControl Then cmdControl

End Sub

