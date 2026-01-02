VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmLrAttribut 
   AutoRedraw      =   -1  'True
   Caption         =   "LR Bafi : mise à jour des attributs"
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9420
   Icon            =   "LrAttribut.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6375
   ScaleWidth      =   9420
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Enregistrer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   0
      Width           =   1200
   End
   Begin VB.CommandButton cmdOption 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   0
      Width           =   1200
   End
   Begin VB.Frame fraOption 
      Caption         =   "Options"
      Height          =   5055
      Left            =   5520
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   3255
      Begin VB.Frame fraLrAttribut 
         Caption         =   "Attribut"
         Height          =   2895
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   3015
         Begin VB.Frame fraFiltre 
            Caption         =   "Filtre"
            Height          =   1575
            Left            =   120
            TabIndex        =   12
            Top             =   1080
            Width           =   2775
            Begin VB.OptionButton optFiltreNE 
               Caption         =   "Différent de"
               Height          =   255
               Left            =   120
               TabIndex        =   16
               Top             =   1080
               Width           =   1215
            End
            Begin VB.OptionButton optFiltreEQ 
               Caption         =   "Egal à"
               Height          =   255
               Left            =   120
               TabIndex        =   15
               Top             =   720
               Value           =   -1  'True
               Width           =   855
            End
            Begin VB.OptionButton optFiltreNO 
               Caption         =   "Ne pas appliquer"
               Height          =   255
               Left            =   120
               TabIndex        =   14
               Top             =   360
               Width           =   1575
            End
            Begin VB.TextBox txtLrAttributValue 
               Height          =   300
               Left            =   1320
               TabIndex        =   13
               Top             =   720
               Width           =   1215
            End
         End
         Begin VB.ComboBox cboLrAttribut 
            Height          =   315
            Left            =   1200
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label lblLrattribut 
            Caption         =   "Sélection "
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   480
            Width           =   975
         End
      End
      Begin VB.Frame fraSort 
         Caption         =   "Trier suivant"
         Height          =   1335
         Left            =   120
         TabIndex        =   6
         Top             =   3480
         Width           =   3015
         Begin VB.OptionButton optSortValue 
            Caption         =   "valeur de l'attribut sélectionné"
            Height          =   375
            Left            =   240
            TabIndex        =   8
            Top             =   840
            Width           =   2415
         End
         Begin VB.OptionButton optSortRéférence 
            Caption         =   "Référence"
            Height          =   375
            Left            =   240
            TabIndex        =   7
            Top             =   360
            Value           =   -1  'True
            Width           =   1215
         End
      End
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   400
      Left            =   8880
      Picture         =   "LrAttribut.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   500
   End
   Begin MSFlexGridLib.MSFlexGrid fgLRAttribut 
      Height          =   5490
      Left            =   480
      TabIndex        =   2
      Top             =   600
      Width           =   8385
      _ExtentX        =   14790
      _ExtentY        =   9684
      _Version        =   393216
      Rows            =   1
      Cols            =   8
      FixedCols       =   0
      RowHeightMin    =   300
      BackColor       =   14737632
      ForeColor       =   12582912
      ForeColorFixed  =   -2147483641
      BackColorSel    =   12648384
      BackColorBkg    =   14737632
      AllowBigSelection=   0   'False
      FocusRect       =   2
      HighLight       =   0
      GridLines       =   0
      GridLinesFixed  =   1
      FormatString    =   "<Nat |<Référence      |<Intitulé                                                   |"
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
      Height          =   400
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   1200
   End
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6240
      TabIndex        =   1
      Top             =   0
      Width           =   2500
   End
   Begin VB.Menu mnuLrAttribut 
      Caption         =   "LrAttribut"
      Visible         =   0   'False
      Begin VB.Menu mnuLRAttributUpdate 
         Caption         =   "Modifier un  enregistrement"
      End
      Begin VB.Menu X1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLRAttributAddNewCompte 
         Caption         =   "Ajouter un compte"
      End
      Begin VB.Menu mnuLRAttributAddNewType 
         Caption         =   "Ajouter un type de compte"
      End
      Begin VB.Menu mnuLRAttributDelete 
         Caption         =   "Supprimer un  enregistrement"
      End
      Begin VB.Menu X2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLRAttributPrint 
         Caption         =   "Impression"
      End
   End
   Begin VB.Menu cmdPrint_mnu 
      Caption         =   "Impression"
      Visible         =   0   'False
      Begin VB.Menu cmdPrint_mnuAll 
         Caption         =   "Imprimer tous les attributs"
      End
      Begin VB.Menu cmdPrint_mnuBafi 
         Caption         =   "Imprimer les attributs Bafi"
      End
      Begin VB.Menu cmdPrint_mnuCdr 
         Caption         =   "Imprimer les attributs Risques"
      End
   End
End
Attribute VB_Name = "frmLrAttribut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim arrTag() As Boolean, arrTagNb As Integer, lstErrClear As Boolean
Dim lastActiveControl_Name  As String, currentActiveControl_Name As String
Dim blnMsgBox_Quit As Boolean, autLrAttribut As typeAuthorization
Dim X As String, I As Integer
Dim Msg As String, valX As String
Dim currentMethod As String, lastMethod As String
Dim blnAddNew As Boolean

Dim recLrAttribut As typeLrAttribut
Dim fgLrAttribut_FormatString As String, fgLrAttribut_K As Integer
Dim fgLrAttribut_BackColorFixed As Long, fgLrAttribut_BackColor As Long

Dim LrAttribut_Name As String, LrAttribut_Value As String
Dim fgLrAttribut_Col As Integer, fgLrAttribut_Colsel As Integer

Dim recAccAut As typeAccAut

Public Sub AccAut_Load()

srvAccAut.Init recAccAut
recAccAut.Method = "SeekP0"
recAccAut.AccAutId = "SRVBIALR"
recAccAut.AccAutK1 = "AUTO"
recAccAut.AccAutK2 = "LRATTRIBUT"
If Not IsNull(srvAccAut.Monitor(recAccAut)) Then Unload Me


If Trim(recAccAut.AccAutTxt) = "" Then
    cmdOk.Enabled = True
    recAccAut.AccAutTxt = usrId
    recAccAut.AccAutDD = DSys
    recAccAut.AccAutHD = time_Hms
    AccAut_Update
Else
    cmdOk.Enabled = False
    MsgBox "Les attributs (LR) sont en cours de mise à jour par : " & Trim(recAccAut.AccAutTxt), vbInformation, "Autorisation : AccAut ( SRVBIALR / AUTO / LRATTRIBUT)"
End If

End Sub


Public Sub AccAut_Update()
recAccAut.Method = constUpdate
If Not IsNull(srvAccAut.Update(recAccAut)) Then
    Call lstErr_AddItem(lstErr, cmdContext, "AccAut : mise à jour non effectuée")
End If

End Sub



Private Sub cboLrAttribut_Click()
LrAttribut_Name = Space$(10)
cbo_Value LrAttribut_Name, cboLrAttribut
End Sub


Private Sub cmdContext_Click()
Select Case cmdContext.Caption
    Case Is = constcmdRechercher
    Case Is = constcmdAbandonner: cmdContext_Quit
End Select

End Sub

Private Sub cmdOk_Click()
'98-10-26 : maj immédiate
'Dim I As Integer
'For I = 1 To arrLrAttributNb
'    If arrLrAttribut(I).Method <> Space(12) And arrLrAttribut(I).Method <> constIgnore Then
'        srvLrAttribut.Update arrLrAttribut(I)
'    End If
'Next I
'blnMsgBox_Quit = False
'cmdContext_Quit
End Sub

Private Sub cmdOption_Click()
If fraOption.Visible Then
    If optSortRéférence Then
        fgLrAttribut_Col = 0: fgLrAttribut_Colsel = 1
    Else
        fgLrAttribut_Col = 3: fgLrAttribut_Colsel = 3
    End If
    fgLrAttribut_Display
    fraOption.Visible = False
    cmdOption.Caption = "&Option"
Else
    fraOption.Visible = True
    cmdOption.Caption = "&Appliquer"
End If
End Sub

'---------------------------------------------------------
Private Sub cmdPrint_Click()
'---------------------------------------------------------
Me.PopupMenu cmdPrint_mnu, vbPopupMenuRightButton
End Sub



Private Sub cmdPrint_mnuAll_Click()
X = Format$(1, "000000") & Format$(arrLrAttributNb, "000000") & " "

prtLrAttributX X

End Sub


Private Sub cmdPrint_mnuBafi_Click()
X = Format$(1, "000000") & Format$(arrLrAttributNb, "000000") & "B"

prtLrAttributX X

End Sub


Private Sub cmdPrint_mnuCdr_Click()
X = Format$(1, "000000") & Format$(arrLrAttributNb, "000000") & "R"

prtLrAttributX X

End Sub


Private Sub fgLrAttribut_Click()
lstErr.Clear
fgLrAttribut_K = fgLRAttribut.Row * fgLRAttribut.Cols
If fgLRAttribut.Row > 0 Then Me.PopupMenu mnuLrAttribut, vbPopupMenuRightButton
End Sub

Private Sub fgLRAttribut_DblClick()
mnuLRAttributUpdate_Click
End Sub


Private Sub fgLrAttribut_GotFocus()
fgLRAttribut.BackColorFixed = focusUsr.BackColor
fgLRAttribut.BackColor = fgLrAttribut_BackColor
End Sub


Private Sub fgLrAttribut_LostFocus()
fgLRAttribut.BackColorFixed = fgLrAttribut_BackColorFixed
'fgLrAttribut.BackColor = vbWindowBackground
End Sub


'---------------------------------------------------------
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'---------------------------------------------------------
Select Case KeyCode
    Case Is = 13: KeyCode = 0: cmdContext_Return
    Case Is = 27: cmdContext_Quit
    Case Is = 44: KeyCode = 0: frmElpPrt.prtScreen
End Select

End Sub


'---------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------
Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)

Call BiaPgmAut_Init("LrAttribut", autLrAttribut)
Form_Clear
ReDim arrLrAttribut(1): arrLrAttributNbMax = 1: arrLrAttributNb = 0
srvLrAttribut.Init recLrAttribut
fgLrAttribut_FormatString = fgLRAttribut.FormatString
fgLrAttribut_BackColorFixed = fgLRAttribut.BackColorFixed
fgLrAttribut_BackColor = fgLRAttribut.BackColor
'If DeviseCoursaut.Consulter Then
    blnAddNew = True
    fgLrAttribut_Load
'End If
fgLrAttribut_Col = 0: fgLrAttribut_Colsel = 1
cboLrAttribut_Load
LrAttribut_Name = cboLrAttribut.List(1)
AccAut_Load

End Sub



'---------------------------------------------------------
Public Sub Form_Clear()
'---------------------------------------------------------
lstErrClear = True
blnMsgBox_Quit = False
usrColor_Set

cmdContext.Enabled = True: cmdContext.BackColor = vbWindowBackground
cmdContext.Caption = constcmdAbandonner: cmdContext.BackColor = errUsr.BackColor
cmdOk.Visible = False
fgLRAttribut.Enabled = True: fgLRAttribut.Clear: fgLRAttribut.Rows = 1
Call lstErr_Clear(lstErr, cmdContext, " 'click' Recherche")
End Sub




Public Sub Msg_Rcv(X As String)
'---------------------------------------------------------

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'If blnMsgBox_Quit Then
'    If MsgBox("Voulez-vous vraiment quitter la saisie sans enregistrer les modifications ?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then Cancel = True
'End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
AccAut_Unload
End Sub

Private Sub mnuLRAttributAddNewCompte_Click()
LrAttributAddNew "C"
End Sub

Private Sub mnuLRAttributAddNewType_Click()
LrAttributAddNew "T"
End Sub

Private Sub LrAttributAddNew(Nature As String)
blnAddNew = True
'fgLRAttribut.Enabled = False
currentMethod = constAddNew
Call lstErr_Clear(lstErr, cmdContext, "Création en cours")
srvLrAttribut.Init arrLrAttribut(0)
arrLrAttribut(0).Nature = Nature
arrLrAttribut(0).Method = constAddNew
frmLrAttributDétail.Display

frmLrAttributDétail_Show

If arrLrAttribut(0).Method = constAddNew Then
'    blnMsgBox_Quit = True
    recLrAttribut = arrLrAttribut(0)
    srvLrAttribut.AddItem recLrAttribut
    arrLrAttribut(arrLrAttributIndex).Method = constAddNew
    srvLrAttribut.Update arrLrAttribut(arrLrAttributIndex)
    fgLrAttribut_Display
End If
End Sub


Public Sub frmLrAttributDétail_Show()
Dim X As String
'frmLrAttributDétail.WindowState = vbNormal
'frmLrAttributDétail.Visible = True

frmLrAttributDétail.Show vbModal 'vbModeless
'X = frmLrAttributDétail.Caption
'AppActivate X

End Sub

Private Sub mnuLRAttributDelete_Click()
fgLrAttribut_Scan
If lstErr.ListCount > 0 Then Exit Sub
'blnMsgBox_Quit = True
'If recLrAttribut.Method = constAddNew Then
'    currentMethod = constIgnore
'    LrAttribut_Delete
'Else
    currentMethod = constDelete
    Call lstErr_Clear(lstErr, cmdContext, "Suppression ligne")
    X = MsgBox("Voulez-vous réellement supprimer cette ligne ?", vbYesNo + vbQuestion + vbDefaultButton2, recLrAttribut.Nature & " " & recLrAttribut.Référence)
    If X = vbYes Then LrAttribut_Delete
    fgLRAttribut.Enabled = True
    fgLRAttribut.SetFocus
'End If
End Sub

Private Sub fgLrAttribut_Load()

Dim blnValidation As Boolean, blnSaisie As Boolean, X As String
srvLrAttribut.Init recLrAttribut
currentMethod = "SnapP0"
recLrAttribut.Method = currentMethod
recLrAttribut.Nature = ""
recLrAttribut.Référence = ""
arrLrAttribut(0) = recLrAttribut
arrLrAttribut(0).Nature = "9"
arrLrAttribut(0).Référence = "9z"
arrLrAttributNb = 0: arrLrAttributIndex = 0
arrLrAttributSuite = True

Do Until Not arrLrAttributSuite
    srvLrAttribut.Monitor recLrAttribut
    recLrAttribut = arrLrAttribut(arrLrAttributNb)
    recLrAttribut.Method = currentMethod & "+"
Loop

fgLrAttribut_Display
Call lstErr_Clear(lstErr, fgLRAttribut, "ok")

End Sub

Private Sub mnuLRAttributUpdate_Click()

fgLrAttribut_Scan
If lstErr.ListCount > 0 Then Exit Sub
'fgLRAttribut.Enabled = False
currentMethod = constUpdate
Call lstErr_Clear(lstErr, cmdContext, "Modification Attributs")

arrLrAttribut(0) = arrLrAttribut(arrLrAttributIndex)
lastMethod = arrLrAttribut(0).Method
arrLrAttribut(0).Method = constUpdate
frmLrAttributDétail.Display

frmLrAttributDétail_Show

If arrLrAttribut(0).Method = constUpdate Then
'    blnMsgBox_Quit = True
    If lastMethod = constAddNew Then arrLrAttribut(0).Method = constAddNew
    arrLrAttribut(arrLrAttributIndex) = arrLrAttribut(0)
    srvLrAttribut.Update arrLrAttribut(arrLrAttributIndex)
    fgLrAttribut_DisplayItem
End If
End Sub



Public Sub fgLrAttribut_Display()
fgLRAttribut.Redraw = False
fgLRAttribut.Clear
fgLRAttribut.Rows = 1
fgLRAttribut.FormatString = fgLrAttribut_FormatString & "<" & LrAttribut_Name
fgLRAttribut.Enabled = True
For arrLrAttributIndex = 1 To arrLrAttributNb
    If arrLrAttribut(arrLrAttributIndex).Method <> constDelete _
    And arrLrAttribut(arrLrAttributIndex).Method <> constIgnore Then
        fgLRAttribut.Rows = fgLRAttribut.Rows + 1
        fgLRAttribut.Row = fgLRAttribut.Rows - 1
        fgLrAttribut_DisplayItem
    End If
Next arrLrAttributIndex
If fgLRAttribut.Rows > 1 Then fgLrAttribut_Sort
fgLRAttribut.Redraw = True

End Sub

Public Sub fgLrAttribut_DisplayItem()
fgLrAttribut_K = (fgLRAttribut.Row) * fgLRAttribut.Cols
fgLRAttribut.TextArray(0 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).Nature
fgLRAttribut.TextArray(1 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).Référence
If arrLrAttribut(arrLrAttributIndex).Nature = "T" Then
    fgLRAttribut.TextArray(2 + fgLrAttribut_K) = DicLib(13, Trim(arrLrAttribut(arrLrAttributIndex).Référence))
Else
    fgLRAttribut.TextArray(2 + fgLrAttribut_K) = "compte ......"
End If

Select Case LrAttribut_Name
    Case "AFFPU": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).AFFPU
    Case "AGEMT": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).AGEMT
    Case "AGENT": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).AGENT
    Case "APPAR": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).APPAR
    Case "AREFR": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).AREFR
    Case "ATTCF": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).ATTCF
    Case "AUTDV": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).AUTDV
    Case "BONIF": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).BONIF
    Case "CAROB": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).CAROB
    Case "CATET": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).CATET
    Case "CDRES": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).CDRES
    Case "CDZON": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).CDZON
    Case "CLCRC": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).CLCRC
    Case "COTIT": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).COTIT
    Case "CPEMS": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).CPEMS
    Case "CRDIV": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).CRDIV
    Case "CREIM": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).CREIM
    Case "CREOR": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).CREOR
    Case "CRETC": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).CRETC
    Case "CRHYP": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).CRHYP
    Case "DCTOM": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).DCTOM
    Case "DRAC": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).DRAC
    Case "DURIN": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).DURIN
    Case "DUROM": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).DUROM
    Case "DVOPR": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).DVOPR
    Case "ECART": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).ECART
    Case "ECFIN": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).ECFIN
    Case "ELIGB": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).ELIGB
    Case "FAMDV": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).FAMDV
    Case "FOPIF": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).FOPIF
    Case "FPRBG": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).FPRBG
    Case "GARCF": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).GARCF
    Case "MLFCE": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).MLFCE
    Case "MONDV": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).MONDV
    Case "MUTFG": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).MUTFG
    Case "NACGA": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).NACGA
    Case "NACGR": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).NACGR
    Case "NACPS": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).NACPS
    Case "NAEGA": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).NAEGA
    Case "NAIMO": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).NAIMO
    Case "NAOCB": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).NAOCB
    Case "NAPRO": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).NAPRO
    Case "NARCP": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).NARCP
    Case "NATCP": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).NATCP
    Case "NATCR": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).NATCR
    Case "NATCS": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).NATCS
    Case "NATDD": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).NATDD
    Case "NATER": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).NATER
    Case "NATIF": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).NATIF
    Case "NATIT": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).NATIT
    Case "NATMA": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).NATMA
    Case "NATOF": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).NATOF
    Case "NATRS": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).NATRS
    Case "NRAST": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).NRAST
    Case "NREHB": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).NREHB
    Case "OPCVM": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).OPCVM
    Case "OPEFC": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).OPEFC
    Case "OPFDH": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).OPFDH
    Case "OPREC": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).OPREC
    Case "PAACT": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).PAACT
    Case "PERIO": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).PERIO
    Case "PRIMP": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).PRIMP
    Case "PROCB": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).PROCB
    Case "REDES": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).REDES
    Case "REDHB": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).REDHB
    Case "RESET": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).RESET
    Case "REZON": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).REZON
    Case "RISPA": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).RISPA
    Case "SENOP": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).SENOP
    Case "TCFPE": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).TCFPE
    Case "TOPIF": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).TOPIF
    Case "TYCGR": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).TYCGR
    Case "TYCOM": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).TYCOM
    Case "TYDSU": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).TYDSU
    Case "TYETS": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).TYETS
    Case "TYPOR": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).TYPOR
    Case "TYPSU": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).TYPSU
    Case "TYRES": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).TYRES
    Case "ACTI": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).ZACTI
    Case "CDCPCO": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).CDCPCO
    Case "CDCPJO": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).CDCPJO
    Case "CDCPFU": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).CDCPFU
    Case "CDAGCO": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).CDAGCO
    Case "CDREME": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).CDREME
    Case "TYMTDV": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).TYMTDV
    Case "TYVENT": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).TYVENT
    Case "CRVENT": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).CRVENT
    Case "CDDURE": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).CDDURE
    Case "DUINIT": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).DUINIT
    Case "CDCRTI": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).CDCRTI
    Case "CDCRAC": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).CDCRAC
    Case "CDBIOR": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).CDBIOR
    Case "CDDEIN": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).CDDEIN
    Case "CDCRIM": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).CDCRIM
    Case "CDCRCO": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).CDCRCO
    Case "CDCREF": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).CDCREF
    Case "CDLODA": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).CDLODA
    Case "CDCRET": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).CDCRET
    Case "CDOMPO": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).CDOMPO
    Case "CDOPIM": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).CDOPIM
    Case "CDSWAP": fgLRAttribut.TextArray(3 + fgLrAttribut_K) = arrLrAttribut(arrLrAttributIndex).CDSWAP

End Select

Select Case arrLrAttribut(arrLrAttributIndex).Method
    Case constAddNew: fgLRAttribut.TextArray(6 + fgLrAttribut_K) = "Créé"
    Case constUpdate: fgLRAttribut.TextArray(6 + fgLrAttribut_K) = "Modifié"
End Select
        
        
If optFiltreEQ Then
    If Trim(fgLRAttribut.TextArray(3 + fgLrAttribut_K)) <> Trim(txtLrAttributValue) Then
        fgLRAttribut.Row = fgLRAttribut.Rows - 2
'        fgLRAttribut.Rows = fgLRAttribut.Rows - 1
    End If
End If
        
If optFiltreNE Then
    If Trim(fgLRAttribut.TextArray(3 + fgLrAttribut_K)) = Trim(txtLrAttributValue) Then
        fgLRAttribut.Row = fgLRAttribut.Rows - 2
'        fgLRAttribut.Rows = fgLRAttribut.Rows - 1
    End If
End If

End Sub

Public Sub LrAttribut_Delete()
'blnMsgBox_Quit = True
recLrAttribut.Method = currentMethod
arrLrAttribut(arrLrAttributIndex) = recLrAttribut
If arrLrAttribut(I).Method <> constIgnore Then srvLrAttribut.Update arrLrAttribut(arrLrAttributIndex)
fgLrAttribut_Display
End Sub

Public Sub cmdContext_Quit()
Dim X As String

If fraOption.Visible Then
    fraOption.Visible = False
    cmdOption.Caption = "&Option"
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
    cmdOption_Click
Else
    SendKeys "{TAB}"
End If

End Sub

Public Sub fgLrAttribut_Scan()
fgLrAttribut_K = fgLRAttribut.Row * fgLRAttribut.Cols
recLrAttribut.Nature = Trim(fgLRAttribut.TextArray(0 + fgLrAttribut_K))
recLrAttribut.Référence = Trim(fgLRAttribut.TextArray(1 + fgLrAttribut_K))
If srvLrAttribut.Scan(recLrAttribut) > 0 Then
    recLrAttribut = arrLrAttribut(arrLrAttributIndex)
Else
    Call lstErr_AddItem(lstErr, fgLRAttribut, "Erreur fgLrAttribut_Scan")
End If
End Sub


Public Sub fgLrAttribut_Sort()
fgLRAttribut.Row = 1
fgLRAttribut.Col = fgLrAttribut_Col

fgLRAttribut.RowSel = 1 'fgLRAttribut.Rows - 1
fgLRAttribut.ColSel = fgLrAttribut_Colsel

fgLRAttribut.Sort = flexSortStringAscending
End Sub





Public Sub cboLrAttribut_Load()
cboLrAttribut.AddItem "AFFPU"
cboLrAttribut.AddItem "AGEMT"
cboLrAttribut.AddItem "AGENT"
cboLrAttribut.AddItem "APPAR"
cboLrAttribut.AddItem "AREFR"
cboLrAttribut.AddItem "ATTCF"
cboLrAttribut.AddItem "AUTDV"
cboLrAttribut.AddItem "BONIF"
cboLrAttribut.AddItem "CAROB"
cboLrAttribut.AddItem "CATET"
cboLrAttribut.AddItem "CDRES"
cboLrAttribut.AddItem "CDZON"
cboLrAttribut.AddItem "CLCRC"
cboLrAttribut.AddItem "COTIT"
cboLrAttribut.AddItem "CPEMS"
cboLrAttribut.AddItem "CRDIV"
cboLrAttribut.AddItem "CREIM"
cboLrAttribut.AddItem "CREOR"
cboLrAttribut.AddItem "CRETC"
cboLrAttribut.AddItem "CRHYP"
cboLrAttribut.AddItem "DCTOM"
cboLrAttribut.AddItem "DRAC"
cboLrAttribut.AddItem "DURIN"
cboLrAttribut.AddItem "DUROM"
cboLrAttribut.AddItem "DVOPR"
cboLrAttribut.AddItem "ECART"
cboLrAttribut.AddItem "ECFIN"
cboLrAttribut.AddItem "ELIGB"
cboLrAttribut.AddItem "FAMDV"
cboLrAttribut.AddItem "FOPIF"
cboLrAttribut.AddItem "FPRBG"
cboLrAttribut.AddItem "GARCF"
cboLrAttribut.AddItem "MLFCE"
cboLrAttribut.AddItem "MONDV"
cboLrAttribut.AddItem "MUTFG"
cboLrAttribut.AddItem "NACGA"
cboLrAttribut.AddItem "NACGR"
cboLrAttribut.AddItem "NACPS"
cboLrAttribut.AddItem "NAEGA"
cboLrAttribut.AddItem "NAIMO"
cboLrAttribut.AddItem "NAOCB"
cboLrAttribut.AddItem "NAPRO"
cboLrAttribut.AddItem "NARCP"
cboLrAttribut.AddItem "NATCP"
cboLrAttribut.AddItem "NATCR"
cboLrAttribut.AddItem "NATCS"
cboLrAttribut.AddItem "NATDD"
cboLrAttribut.AddItem "NATER"
cboLrAttribut.AddItem "NATIF"
cboLrAttribut.AddItem "NATIT"
cboLrAttribut.AddItem "NATMA"
cboLrAttribut.AddItem "NATOF"
cboLrAttribut.AddItem "NATRS"
cboLrAttribut.AddItem "NRAST"
cboLrAttribut.AddItem "NREHB"
cboLrAttribut.AddItem "OPCVM"
cboLrAttribut.AddItem "OPEFC"
cboLrAttribut.AddItem "OPFDH"
cboLrAttribut.AddItem "OPREC"
cboLrAttribut.AddItem "PAACT"
cboLrAttribut.AddItem "PERIO"
cboLrAttribut.AddItem "PRIMP"
cboLrAttribut.AddItem "PROCB"
cboLrAttribut.AddItem "REDES"
cboLrAttribut.AddItem "REDHB"
cboLrAttribut.AddItem "RESET"
cboLrAttribut.AddItem "REZON"
cboLrAttribut.AddItem "RISPA"
cboLrAttribut.AddItem "SEMNT"
cboLrAttribut.AddItem "SENOP"
cboLrAttribut.AddItem "TCFPE"
cboLrAttribut.AddItem "TOPIF"
cboLrAttribut.AddItem "TYCGR"
cboLrAttribut.AddItem "TYCOM"
cboLrAttribut.AddItem "TYDSU"
cboLrAttribut.AddItem "TYETS"
cboLrAttribut.AddItem "TYPOR"
cboLrAttribut.AddItem "TYPSU"
cboLrAttribut.AddItem "TYRES"
cboLrAttribut.AddItem "ZACTI"

'attibuts Luca Risques
cboLrAttribut.AddItem "CDCPCO"
cboLrAttribut.AddItem "CDCPJO"
cboLrAttribut.AddItem "CDCPFU"
cboLrAttribut.AddItem "CDAGCO"
cboLrAttribut.AddItem "CDREME"
cboLrAttribut.AddItem "TYMTDV"
cboLrAttribut.AddItem "TYVENT"
cboLrAttribut.AddItem "CRVENT"
cboLrAttribut.AddItem "CDDURE"
cboLrAttribut.AddItem "DUINIT"
cboLrAttribut.AddItem "CDCRTI"
cboLrAttribut.AddItem "CDCRAC"
cboLrAttribut.AddItem "CDBIOR"
cboLrAttribut.AddItem "CDDEIN"
cboLrAttribut.AddItem "CDCRIM"
cboLrAttribut.AddItem "CDCRCO"
cboLrAttribut.AddItem "CDCREF"
cboLrAttribut.AddItem "CDLODA"
cboLrAttribut.AddItem "CDCRET"
cboLrAttribut.AddItem "CDOMPO"
cboLrAttribut.AddItem "CDOPIM"
cboLrAttribut.AddItem "CDSWAP"

End Sub


Public Sub AccAut_Unload()
If cmdOk.Enabled Then
    recAccAut.AccAutTxt = ""
    recAccAut.AccAutDF = DSys
    recAccAut.AccAutHF = time_Hms
    AccAut_Update
End If

End Sub
