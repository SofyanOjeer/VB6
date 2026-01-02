VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBiaLog 
   AutoRedraw      =   -1  'True
   Caption         =   "BiaLog : journal des événements"
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   6255
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   11033
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Sélection"
      TabPicture(0)   =   "BiaLog.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraOption"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Liste des événements"
      TabPicture(1)   =   "BiaLog.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fgSelect"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Détail"
      TabPicture(2)   =   "BiaLog.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
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
         Begin VB.ComboBox cboSelectCodeErr 
            Height          =   315
            Left            =   2760
            TabIndex        =   20
            Text            =   "Combo1"
            Top             =   3240
            Width           =   5775
         End
         Begin VB.CheckBox chkSelectCodeErr 
            Caption         =   "Code Erreur"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   3240
            Width           =   2055
         End
         Begin VB.TextBox txtSelectDossierMax 
            Height          =   285
            Left            =   4800
            MaxLength       =   11
            TabIndex        =   18
            Top             =   2520
            Width           =   1575
         End
         Begin VB.TextBox txtSelectDossierMin 
            Height          =   285
            Left            =   2760
            MaxLength       =   11
            TabIndex        =   17
            Top             =   2520
            Width           =   1575
         End
         Begin VB.CheckBox chkSelectDossier 
            Caption         =   "sélectionner les dossiers"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   2520
            Width           =   2415
         End
         Begin VB.CheckBox chkSelectAmj 
            Caption         =   "Période"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   1200
            Width           =   2535
         End
         Begin VB.CheckBox chkSelectService 
            Caption         =   "Service"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   600
            Width           =   2535
         End
         Begin VB.TextBox txtSelectCompteMax 
            Height          =   285
            Left            =   4800
            MaxLength       =   11
            TabIndex        =   12
            Top             =   1800
            Width           =   1575
         End
         Begin VB.TextBox txtSelectCompteMin 
            Height          =   285
            Left            =   2760
            MaxLength       =   11
            TabIndex        =   11
            Top             =   1800
            Width           =   1575
         End
         Begin VB.CheckBox chkSelectCompte 
            Caption         =   "sélectionner les comptes "
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   1800
            Width           =   2295
         End
         Begin VB.CommandButton cmdSelect 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Rechercher"
            Height          =   975
            Left            =   2880
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   3960
            Width           =   2415
         End
         Begin VB.TextBox txtSelectService 
            Height          =   285
            Left            =   2760
            TabIndex        =   6
            Top             =   600
            Width           =   975
         End
         Begin MSComCtl2.DTPicker txtSelectAmjMax 
            Height          =   300
            Left            =   4800
            TabIndex        =   7
            Top             =   1200
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
            Format          =   65077251
            CurrentDate     =   36299
            MaxDate         =   401768
            MinDate         =   -328351
         End
         Begin MSComCtl2.DTPicker txtSelectAmjMin 
            Height          =   300
            Left            =   2760
            TabIndex        =   15
            Top             =   1200
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
            Format          =   65077251
            CurrentDate     =   36299
            MaxDate         =   401768
            MinDate         =   -328351
         End
      End
      Begin MSFlexGridLib.MSFlexGrid fgSelect 
         Height          =   5250
         Left            =   -74880
         TabIndex        =   4
         Top             =   600
         Width           =   9075
         _ExtentX        =   16007
         _ExtentY        =   9260
         _Version        =   393216
         Rows            =   1
         Cols            =   13
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
         FormatString    =   $"BiaLog.frx":0054
      End
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   500
      Left            =   8880
      Picture         =   "BiaLog.frx":0192
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
   Begin VB.Menu mnuOpération 
      Caption         =   "Opération"
      Visible         =   0   'False
      Begin VB.Menu mnuOpérationDisplay 
         Caption         =   "Afficher ce contrat"
      End
   End
   Begin VB.Menu mnuPrint 
      Caption         =   "Print"
      Visible         =   0   'False
      Begin VB.Menu mnuOpérationPrint 
         Caption         =   "Imprimer ce contrat"
      End
      Begin VB.Menu mnuListPrint 
         Caption         =   "Imprimer la liste"
      End
   End
End
Attribute VB_Name = "frmBiaLog"
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
Dim BiaLogAut As typeAuthorization

Dim fgSelect_FormatString As String, fgSelect_K As Integer
Dim fgSelect_RowDisplay As Integer, fgSelect_RowClick As Integer
Dim fgSelect_ColorClick As Long, fgSelect_ColorDisplay As Long
Dim fgSelect_Sort1 As Integer, fgSelect_Sort2 As Integer
Dim fgSelect_SortAD As Integer, fgSelect_Sort1_Old As Integer
Dim fgSelect_arrIndex As Integer
Dim recBiaLog As typeBiaLog, xBiaLog As typeBiaLog, mBiaLog As typeBiaLog, mEchéancierBiaLog As typeBiaLog

Dim meBiaLog() As typeBiaLog
Dim meBiaLog_Nb As Integer, meBiaLog_Index As Integer, meBiaLog_NbMax As Integer

Dim blncmdOk_Visible As Boolean, blnErr As Boolean, blncmdSave_Visible As Boolean
Dim blnfgSelect_DisplayLine As Boolean, blnfgEchéance_DisplayLine As Boolean


Dim blnSetfocus As Boolean
Dim blnSelectCompte As Boolean, wSelectCompteMin As String * 11, wSelectCompteMax As String * 11
Dim blnSelectService As Boolean, wSelectService As String * 5, blnSelectService_Enabled As Boolean
Dim blnSelectDossier As Boolean, wSelectDossierMin, wSelectDossierMax
Dim blnSelectAmj As Boolean, wSelectAmjMin As String * 8, wSelectAmjMax As String * 8
Dim blnSelectCodeErr As Boolean, wSelectCodeErr As String * 12

Dim recCompte As typeCompte

Private Sub chkSelectAmj_Click()
On Error GoTo Exit_Sub
If chkSelectAmj = "1" Then
    txtSelectAmjMin.Visible = True: txtSelectAmjMax.Visible = True
    If blnSetfocus Then txtSelectAmjMin.SetFocus
Else
    txtSelectAmjMin.Visible = False: txtSelectAmjMax.Visible = False
End If
Exit_Sub:

End Sub

Private Sub chkSelectAmj_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkSelectAmj
End Sub


Private Sub chkSelectCodeErr_Click()
On Error GoTo Exit_Sub

If chkSelectCodeErr = "1" Then
    cboSelectCodeErr.Visible = True
    If blnSetfocus Then cboSelectCodeErr.SetFocus
Else
    cboSelectCodeErr.Visible = False
    cboSelectCodeErr.ListIndex = 0
End If
Exit_Sub:

End Sub

Private Sub chkSelectCompte_Click()

On Error GoTo Exit_Sub

If chkSelectCompte = "1" Then
    txtSelectCompteMin.Visible = True: txtSelectCompteMax.Visible = True
    If blnSetfocus Then txtSelectCompteMin.SetFocus
Else
    txtSelectCompteMin.Visible = False: txtSelectCompteMax.Visible = False
End If
Exit_Sub:

End Sub

Private Sub chkSelectCompte_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkSelectCompte
End Sub

'---------------------------------------------------------
Public Sub arrTag_Set(ByVal B As Boolean)
'---------------------------------------------------------
For I = 0 To arrTagNb: arrTag(I) = B: Next I
End Sub


Public Sub cmdContext_Quit()
blnControl = False

lstErr.Clear
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
wSelectService = Format$(Val(Trim(txtSelectService)), "000") & "  "
If blnSelectService Then
    If wSelectService = "00000" Then Call lstErr_AddItem(lstErr, cmdContext, "? préciser le service")
End If

blnSelectCompte = IIf(chkSelectCompte = "1", True, False)
wSelectCompteMin = Format$(Val(Trim(txtSelectCompteMin)), "00000000000")
wSelectCompteMax = Format$(Val(Trim(txtSelectCompteMax)), "00000000000")
If blnSelectCompte Then
    If wSelectCompteMin = "00000000000" Then
        Call lstErr_AddItem(lstErr, cmdContext, "? préciser le compte min")
    Else
        If wSelectCompteMax = "00000000000" Then wSelectCompteMax = wSelectCompteMin
    End If
    If wSelectCompteMin > wSelectCompteMax Then Call lstErr_AddItem(lstErr, cmdContext, "? compte min > compte max")
End If

blnSelectDossier = IIf(chkSelectDossier = "1", True, False)
wSelectDossierMin = Trim(txtSelectDossierMin)
wSelectDossierMax = Trim(txtSelectDossierMax)
If blnSelectDossier Then
    If wSelectDossierMin = "" Then
        Call lstErr_AddItem(lstErr, cmdContext, "? préciser le dossier min")
    Else
        If wSelectDossierMax = "" Then wSelectDossierMax = wSelectDossierMin
    End If
    If wSelectDossierMin > wSelectDossierMax Then Call lstErr_AddItem(lstErr, cmdContext, "? dossier min > dossier max")
End If

blnSelectAmj = IIf(chkSelectAmj = "1", True, False)
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

blnSelectCodeErr = IIf(chkSelectCodeErr = "1", True, False)
cbo_Value wSelectCodeErr, cboSelectCodeErr
If blnSelectCodeErr Then
    If Trim(wSelectCodeErr) = "" Then Call lstErr_AddItem(lstErr, cmdContext, "? préciser le code erreur")
End If



'cmdSave.Visible = blncmdSave_Visible
If lstErr.ListCount > 0 Then
    lstErr.Visible = True
Else
    'cmdOk.Visible = blncmdOk_Visible
    'blnSetfocus = True: currentActiveControl_Name = "cmdOk"
End If

ExitSub:

Me.Enabled = True
    
blnControl = True

End Sub

Private Sub chkSelectDossier_Click()
On Error GoTo Exit_Sub
If chkSelectDossier = "1" Then
    txtSelectDossierMin.Visible = True: txtSelectDossierMax.Visible = True
    If blnSetfocus Then txtSelectDossierMin.SetFocus
Else
    txtSelectDossierMin.Visible = False: txtSelectDossierMax.Visible = False
End If
Exit_Sub:

End Sub

Private Sub chkSelectDossier_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set chkSelectDossier

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


Private Sub txtSelectCompteMax_GotFocus()

txt_GotFocus txtSelectCompteMax

End Sub


Private Sub txtSelectCompteMax_KeyPress(KeyAscii As Integer)
Call num_KeyAscii(KeyAscii)

End Sub


Private Sub txtSelectCompteMax_LostFocus()
txt_LostFocus txtSelectCompteMax

End Sub

Private Sub txtSelectCompteMin_GotFocus()
txt_GotFocus txtSelectCompteMin

End Sub

Private Sub txtSelectCompteMin_KeyPress(KeyAscii As Integer)
Call num_KeyAscii(KeyAscii)


End Sub


Private Sub txtSelectCompteMin_LostFocus()
txt_LostFocus txtSelectCompteMin


End Sub

Private Sub txtSelectDossierMax_GotFocus()
txt_GotFocus txtSelectDossierMax


End Sub


Private Sub txtSelectDossierMax_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtSelectDossierMax_LostFocus()
txt_LostFocus txtSelectDossierMax

End Sub


Private Sub txtSelectDossierMin_GotFocus()
txt_GotFocus txtSelectDossierMin


End Sub


Private Sub txtSelectDossierMin_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtSelectDossierMin_LostFocus()
txt_LostFocus txtSelectDossierMin

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

cmdContext.Caption = constcmdRechercher: blnMsgBox_Quit = False
arrTag_Set False
lstErr.Visible = False
blncmdOk_Visible = False: blncmdSave_Visible = False
fgSelect_Reset


chkSelectAmj.Value = "1":    txtSelectAmjMin.Visible = True: txtSelectAmjMax.Visible = True
Call DTPicker_Set(txtSelectAmjMin, AmjCptVeille)
Call DTPicker_Set(txtSelectAmjMax, DSys)

chkSelectCompte.Value = "0":    txtSelectCompteMin.Visible = False: txtSelectCompteMax.Visible = False
chkSelectDossier.Value = "0":   txtSelectDossierMin.Visible = False: txtSelectDossierMax.Visible = False
chkSelectCodeErr.Value = "0":   cboSelectCodeErr.Visible = False: wSelectCodeErr = ""

Call cboDictio(524, cboSelectCodeErr, 12)
If cboSelectCodeErr.ListCount > 0 Then cboSelectCodeErr.ListIndex = 0

wSelectService = "00000"
If usrService_DisplayAll Then
    chkSelectService.Value = "0": txtSelectService = "": txtSelectService.Visible = False
Else
    txtSelectService = usrService: txtSelectService.Visible = True
    chkSelectService.Value = "1"
    chkSelectService.Enabled = False: txtSelectService.Enabled = False
End If
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

Public Function Compte_Load(lDevise As String, lCompteNuméro As String)
Compte_Load = Null
recCompte.Devise = lDevise
recCompte.Numéro = lCompteNuméro
If blnRéplication_Load Then
    V = mdbCptP0_Find(recCompte)
Else
    V = "Compte_Load à revoir "  'srvCompteFind(recCompte)
End If

If Not IsNull(V) Then Call lstErr_AddItem(lstErr, lstErr, "? compte inconnu : " & lDevise & lCompteNuméro): Compte_Load = "?": Exit Function

If recCompte.Situation <> " " Then
    Select Case recCompte.Situation
        Case "B": Call lstErr_AddItem(lstErr, lstErr, " ? Compte bloqué : " & lCompteNuméro): Compte_Load = "?"
        Case "A": Call lstErr_AddItem(lstErr, lstErr, " ? Compte annulé : " & lCompteNuméro): Compte_Load = "?"
        Case Else: Call lstErr_AddItem(lstErr, lstErr, " ? Situation du compte : " & lCompteNuméro): Compte_Load = "?"
    End Select
End If

End Function
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

SSTab1.Tab = 1

fgSelect.Visible = True
fgSelect.Clear: fgSelect.Rows = 1: fgSelect_RowDisplay = 0

fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
fgSelect.Enabled = True
For meBiaLog_Index = 1 To meBiaLog_Nb
    If meBiaLog(meBiaLog_Index).Method <> constIgnore And meBiaLog(meBiaLog_Index).Method <> constDelete Then
        fgSelect.Rows = fgSelect.Rows + 1
        fgSelect.Row = fgSelect.Rows - 1
        fgSelect_DisplayLine
    End If
Next meBiaLog_Index

fgSelect_SortAD = 5
If fgSelect.Rows > 1 Then fgSelect_Sort

End Sub
Public Sub fgSelect_DisplayLine()

fgSelect.Col = 0: fgSelect.Text = meBiaLog(meBiaLog_Index).Log_CodErr
fgSelect.Col = 1: fgSelect.Text = Compte_Imp(meBiaLog(meBiaLog_Index).Log_Compte)
fgSelect.Col = 2: fgSelect.Text = meBiaLog(meBiaLog_Index).Log_RefCon
fgSelect.Col = 3: fgSelect.Text = meBiaLog(meBiaLog_Index).Log_Texte1
fgSelect.Col = 4: fgSelect.Text = meBiaLog(meBiaLog_Index).Log_Texte2
fgSelect.Col = 5: fgSelect.Text = meBiaLog(meBiaLog_Index).Log_Servic
fgSelect.Col = 6: fgSelect.Text = meBiaLog(meBiaLog_Index).Log_Profil
fgSelect.Col = 7: fgSelect.Text = meBiaLog(meBiaLog_Index).Log_Progr
fgSelect.Col = 8: fgSelect.Text = dateImp(meBiaLog(meBiaLog_Index).Log_CptAmj)
fgSelect.Col = 9: fgSelect.Text = Format(meBiaLog(meBiaLog_Index).Log_Cpteur, "#### ### ##0 ")
fgSelect.Col = 10: fgSelect.Text = dateImp(meBiaLog(meBiaLog_Index).Log_SysAmj) & timeImp(meBiaLog(meBiaLog_Index).Log_SysAmj)

fgSelect.Col = fgSelect_arrIndex - 1: fgSelect.Text = ""
fgSelect.Col = fgSelect_arrIndex: fgSelect.Text = meBiaLog_Index

End Sub
Public Sub fgSelect_Load()
Dim X As String, mMethod As String

recBiaLog_Init xBiaLog

xBiaLog.Method = "SnapL2"
xBiaLog.Log_Cosoc = SocId$
xBiaLog.Log_Agence = SocAgence$
xBiaLog.Log_CptAmj = "00000000"
xBiaLog.Log_Servic = wSelectService

meBiaLog(0) = xBiaLog
meBiaLog(0).Log_CptAmj = "99999999"

If Not blnSelectService Then
    xBiaLog.Log_Servic = "00000"
    meBiaLog(0).Log_Servic = "99999"
End If

If blnSelectAmj Then
    If Not blnSelectService Then xBiaLog.Method = "SnapP0"
    xBiaLog.Log_CptAmj = wSelectAmjMin
    meBiaLog(0).Log_CptAmj = wSelectAmjMax
End If


If blnSelectCompte Then
    xBiaLog.Method = "SnapL4"
    xBiaLog.Log_Compte = wSelectCompteMin
    meBiaLog(0).Log_Compte = wSelectCompteMax
End If

If blnSelectDossier Then
    xBiaLog.Method = "SnapL3"
    xBiaLog.Log_RefCon = wSelectDossierMin
    meBiaLog(0).Log_RefCon = wSelectDossierMax
End If


If blnSelectCodeErr Then
    If Not blnSelectService Then xBiaLog.Method = "SnapL5"
    xBiaLog.Log_CodErr = wSelectCodeErr
    meBiaLog(0).Log_CodErr = wSelectCodeErr
End If

Call srvBiaLog_Load(xBiaLog, meBiaLog(0))

meBiaLog_Nb = srvBiaLog.arrBiaLog_NB
meBiaLog_NbMax = meBiaLog_Nb + 1: ReDim meBiaLog(meBiaLog_NbMax)

For I = 1 To meBiaLog_Nb
    meBiaLog(I) = srvBiaLog.arrBiaLog(I)
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
    meBiaLog_Index = Val(fgSelect.Text)
    fgSelect.Col = fgSelect_arrIndex - 1
    X = meBiaLog(meBiaLog_Index).Log_CodErr & meBiaLog(meBiaLog_Index).Log_Compte & meBiaLog(meBiaLog_Index).Log_RefCon
    Select Case lK
        Case 1: fgSelect.Text = meBiaLog(meBiaLog_Index).Log_Compte & X
        Case 5: fgSelect.Text = meBiaLog(meBiaLog_Index).Log_Servic & X
        Case 6: fgSelect.Text = meBiaLog(meBiaLog_Index).Log_Profil & X
        Case 7: fgSelect.Text = meBiaLog(meBiaLog_Index).Log_Servic & X
        Case 8: fgSelect.Text = meBiaLog(meBiaLog_Index).Log_Progr & X
        Case 10: fgSelect.Text = meBiaLog(meBiaLog_Index).Log_SysAmj & X
        Case fgSelect_arrIndex: fgSelect.Text = Format$(meBiaLog_Index, "0000000000")
    End Select
Next I

fgSelect_Sort1 = fgSelect_arrIndex - 1: fgSelect_Sort2 = fgSelect_arrIndex - 1
fgSelect_Sort
End Sub


Public Sub Form_Init()
Call lstErr_Clear(lstErr, cmdContext, "Initialisation ")
DoEvents

SSTab1.Tab = 0
ReDim meBiaLog(10)

blnControl = False
fgSelect_FormatString = fgSelect.FormatString

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
Dim Msg As String
Dim I As Integer

Me.Enabled = False

Msg = Space$(50)
prtBiaLog_Open Msg

For I = 1 To fgSelect.Rows - 1
    fgSelect.Row = I
    fgSelect.Col = fgSelect_arrIndex
    meBiaLog_Index = Val(fgSelect.Text)
    fgSelect.Col = fgSelect_arrIndex - 1
    recBiaLog = meBiaLog(meBiaLog_Index)
    prtBiaLog_Line recBiaLog
Next I


prtBiaLog_Close

Me.Enabled = True

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
        Case 0: fgSelect_Sort1 = 0: fgSelect_Sort2 = 2: fgSelect_Sort
        Case 1: fgSelect_SortX 1
        Case 2: fgSelect_Sort1 = 2: fgSelect_Sort2 = 2: fgSelect_Sort
        Case 3: fgSelect_Sort1 = 3: fgSelect_Sort2 = 3: fgSelect_Sort
        Case 4: fgSelect_Sort1 = 4: fgSelect_Sort2 = 4: fgSelect_Sort
        Case 5: fgSelect_SortX 5
        Case 6: fgSelect_SortX 6
        Case 7: fgSelect_SortX 7
        Case 8: fgSelect_SortX 8
        Case 9: fgSelect_Sort1 = 9: fgSelect_Sort2 = 9: fgSelect_Sort
        Case 10: fgSelect_SortX 10
        Case fgSelect_arrIndex:  fgSelect_SortX fgSelect_arrIndex
    End Select
Else
    If fgSelect.Rows > 1 Then
        fgSelect.Col = fgSelect_arrIndex
        meBiaLog_Index = Val(fgSelect.Text)
        mBiaLog = meBiaLog(meBiaLog_Index)
        Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
    
        mnuOpérationDisplay = BiaLogAut.Consulter
    
        Me.PopupMenu mnuOpération, vbPopupMenuLeftButton
    End If
End If

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

Private Sub mnuOpérationDisplay_Click()
srvBiaLog_ElpDisplay mBiaLog
End Sub

Public Sub Msg_Rcv(Msg As String)
'---------------------------------------------------------
mWindowState = Me.WindowState
If Me.WindowState <> frmUsr_Windowstate Then Me.WindowState = frmUsr_Windowstate

Call BiaPgmAut_Init(mId$(Msg, 1, 12), BiaLogAut)    ' "SOBF_Effets"

blnSetfocus = True
Form_Init

If UCase$(Trim(mId$(Msg, 13, 12))) = "BIA_EXPLOIT" Then
    cmdSelect_Click
    cmdPrint_Click
    Unload Me
End If

End Sub


Public Sub cmdContext_Return()
If SSTab1.Tab > 0 Then
    SSTab1.Tab = 0
Else
    SendKeys "{TAB}"
    
End If

End Sub



Private Sub txtAmjMax_GotFocus()
'txt_GotFocus txtSelect

End Sub

Private Sub txtAmjMax_LostFocus()

End Sub



Public Sub fgSelect_Reset()
fgSelect_Sort1 = 1: fgSelect_Sort2 = 2
fgSelect_Sort1_Old = 0
fgSelect_RowDisplay = 0: fgSelect_RowClick = 0
fgSelect_arrIndex = 12
blnfgSelect_DisplayLine = False

End Sub
