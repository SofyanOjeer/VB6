VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmElpDoc 
   Caption         =   "Documentation"
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9420
   LinkTopic       =   "Form1"
   ScaleHeight     =   6375
   ScaleWidth      =   9420
   Begin VB.CommandButton cmdServiceX 
      BackColor       =   &H00C0C0FF&
      Caption         =   "X_Service"
      Height          =   435
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   0
      Width           =   900
   End
   Begin VB.CommandButton cmdAddNew 
      Caption         =   "Ajouter un &Dossier"
      Height          =   495
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton cmdContext 
      Caption         =   "&Abandon"
      Height          =   495
      Left            =   0
      Picture         =   "ElpDoc.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   400
      Left            =   8940
      Picture         =   "ElpDoc.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   500
   End
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4680
      TabIndex        =   1
      Top             =   0
      Width           =   4305
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   5775
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   10186
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Recherche"
      TabPicture(0)   =   "ElpDoc.frx":0544
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraFilter"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraSelect"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Détail"
      TabPicture(1)   =   "ElpDoc.frx":0560
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraDétail"
      Tab(1).Control(1)=   "cmdOk"
      Tab(1).Control(2)=   "fgDoc"
      Tab(1).ControlCount=   3
      Begin VB.Frame fraSelect 
         Caption         =   "Sélection"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   5175
         Left            =   3360
         TabIndex        =   14
         Top             =   480
         Width           =   5775
         Begin VB.CommandButton cmdPrior 
            Caption         =   "&Précédent"
            Height          =   1215
            Left            =   4200
            Picture         =   "ElpDoc.frx":057C
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   240
            Width           =   1335
         End
         Begin VB.ListBox lstSelect 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00800000&
            Height          =   3375
            Left            =   240
            TabIndex        =   16
            Top             =   1560
            Width           =   5295
         End
         Begin VB.ListBox lstSelectFilter 
            BackColor       =   &H00808080&
            Height          =   1230
            Left            =   240
            TabIndex        =   15
            Top             =   240
            Width           =   3855
         End
      End
      Begin VB.Frame fraFilter 
         Caption         =   "Filtre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5175
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   3015
         Begin VB.ListBox lstSelectTable 
            Height          =   3375
            Left            =   240
            TabIndex        =   12
            Top             =   1560
            Width           =   2535
         End
         Begin VB.ListBox lstSelectPlan 
            BackColor       =   &H00FFFFFF&
            Height          =   1230
            Left            =   240
            TabIndex        =   11
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.Frame fraDétail 
         Height          =   5175
         Left            =   -74760
         TabIndex        =   4
         Top             =   480
         Width           =   2895
         Begin VB.FileListBox filDoc 
            Height          =   2040
            Left            =   600
            TabIndex        =   9
            Top             =   3000
            Width           =   2535
         End
         Begin VB.ListBox lstTable 
            Height          =   2010
            Left            =   120
            TabIndex        =   7
            Top             =   3120
            Width           =   2535
         End
         Begin VB.ListBox lstPlan 
            Height          =   2600
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   2535
         End
         Begin VB.TextBox txtMemo 
            Height          =   375
            Left            =   240
            TabIndex        =   5
            Top             =   2640
            Width           =   2535
         End
      End
      Begin VB.CommandButton cmdOk 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Ok"
         Height          =   525
         Left            =   -66600
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   50
         Width           =   660
      End
      Begin MSFlexGridLib.MSFlexGrid fgDoc 
         Height          =   5130
         Left            =   -71640
         TabIndex        =   8
         Top             =   600
         Width           =   5835
         _ExtentX        =   10292
         _ExtentY        =   9049
         _Version        =   393216
         Rows            =   1
         Cols            =   5
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
         FormatString    =   "Plan                   |<  Libellé                                                                               |||"
      End
   End
   Begin VB.Menu mnu 
      Caption         =   "mnu"
      Visible         =   0   'False
      Begin VB.Menu mnuDoc_Display 
         Caption         =   "Consulter le document"
      End
      Begin VB.Menu mnuDossier_Display 
         Caption         =   "Consulter le dossier"
      End
      Begin VB.Menu mnuX1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDoc_Print 
         Caption         =   "imprimer le document"
      End
      Begin VB.Menu mnuDossier_Print 
         Caption         =   "Imprimer la page de présentation"
      End
      Begin VB.Menu mnuElpDoc_Print 
         Caption         =   "imprimer la fiche technique"
      End
      Begin VB.Menu mnuSelect_Print 
         Caption         =   "imprimer la sélection"
      End
      Begin VB.Menu mnuX2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDossier_Addnew 
         Caption         =   "Ajouter un dossier"
      End
      Begin VB.Menu mnuDossier_Update 
         Caption         =   "Modifier un dossier"
      End
      Begin VB.Menu mnuDossier_Delete 
         Caption         =   "Supprimer un dossier"
      End
      Begin VB.Menu mnuX3 
         Caption         =   "-"
      End
      Begin VB.Menu mnufgDoc_Modifier 
         Caption         =   "Modifier une ligne"
      End
      Begin VB.Menu mnufgDoc_Supprimer 
         Caption         =   "Supprimer une ligne"
      End
   End
End
Attribute VB_Name = "frmElpDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arrTag() As Boolean, arrTagNb As Integer, lstErrClear As Boolean
Dim lastActiveControl_Name  As String, currentActiveControl_Name As String
Dim MouseMoveActiveControl_Name  As String, MouseMoveActiveControl As typeUsrColor
Dim currentMethod As String

Dim blnMsgBox_Quit As Boolean
Dim DocAut As typeAuthorization, blnInput As Boolean
Dim Msg As String, lstI As Integer

Dim paramElpDoc_Archive As String
Dim paramElpDoc_Validation As String

Dim arrPlan() As typeElpTable, lnkPlan() As typeElpTable
Dim arrTable() As typeElpTable, lnkTable() As typeElpTable
Dim arrDoc() As typeElpDoc, recDoc As typeElpDoc, arrDoc_Index As Integer
Dim blnAddNew As Boolean, blnAddNew_Document As Boolean, blnAddNew_Service As Boolean

Dim fgDoc_FormatString As String, fgDoc_K As Integer
Dim fgDoc_BackColorFixed As Long, fgDoc_BackColor As Long
Dim mLstPlan_ListIndex As Integer
Dim arrSelect() As typeElpDoc
Dim arrSelectPlan() As typeElpTable, arrSelectTable() As typeElpTable
Dim arrSelectSN() As String * 12, arrSelectSN_Nb As Integer, arrSelectSN_Min As Integer
Dim arrSelectSNFilter() As Integer

Dim usrService_Min As String, usrService_Max As String
Dim xElpDoc As typeElpDoc

Public Sub cmdContext_Quit()
If SSTab.Tab = 1 Then
    ssTab1_Clear
    SSTab.Tab = 0
Else
    If lstSelectFilter.ListCount > 0 Then
        form_Clear
    Else
        Unload Me
    End If
End If

End Sub
Public Sub form_Clear()
lstErr.Clear
lstSelect.Clear: arrSelectSN_Nb = 0
lstSelectTable.Clear
lstSelectFilter.Clear
lstSelectFilter.BackColor = &HB0B0B0
lstSelectFilter.ForeColor = &HFFFFFF
'lstSelect.BackColor = &HB0B0B0
SSTab.Tab = 0
ssTab1_Clear
cmdPrior.Visible = False

usrService_Min = "": usrService_Max = ""
recElpTable_Init recElpTable
recElpTable.Id = "Doc_Annuaire"
recElpTable.K1 = usrId
recElpTable.Method = "Seek="
tableElpTable_Read recElpTable
If Trim(recElpTable.Err) = "" Then
    usrService_Min = mId$(recElpTable.Memo, 1, 3)
    usrService_Max = usrService_Min
    If mId$(usrService_Max, 3, 1) = "0" Then
        Mid$(usrService_Max, 3, 1) = "9"
        If mId$(usrService_Max, 2, 1) = "0" Then Mid$(usrService_Max, 2, 1) = "9"
    End If
End If

cmdAddNew.Visible = DocAut.Saisir
If DocAut.Saisir Then Call lstErr_Clear(lstErr, cmdContext, "DocAut.saisir => pas de contrôle de confidentialité")
End Sub
'-------------------------------------------------------
Sub txt_GotFocus(C As Control)
'-------------------------------------------------------
lstErr.Clear
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




Public Sub Msg_Rcv(txtMsg As String)

Call BiaPgmAut_Init(txtMsg, DocAut)
cmdServiceX.Visible = DocAut.Xspécial

If DocAut.Saisir Then MDB_Master

paramElpDoc_Validation = ""
recElpTable_Init recElpTable
recElpTable.Id = "Doc_Param"
recElpTable.K1 = "Validation"
recElpTable.K2 = ""
recElpTable.Method = "Seek="
If Not IsNull(dbElpTable_ReadE(recElpTable)) Then Unload Me
paramElpDoc_Validation = paramServer(recElpTable.Memo)

paramElpDoc_Archive = ""
recElpTable_Init recElpTable
recElpTable.Id = "Doc_Param"
recElpTable.K1 = "Archive"
recElpTable.K2 = ""
recElpTable.Method = "Seek="
If Not IsNull(dbElpTable_ReadE(recElpTable)) Then Unload Me
paramElpDoc_Archive = paramServer(recElpTable.Memo)

form_Clear

End Sub

Private Sub cmdAddNew_Click()
mnuDossier_Addnew_Click
End Sub

Private Sub cmdContext_Click()
cmdContext_Quit
End Sub

Private Sub cmdContext_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdContext
End Sub


Private Sub cmdOk_Click()
Dim X As String, xDest As String, xSource As String
On Error GoTo Error_Msg

X = Trim(CStr(arrDoc(0).Memo))
xDest = Trim(paramServer(lnkTable(1).Memo)) & "\" & X
xSource = paramElpDoc_Validation & "\" & X

If msFileSystem.FileExists(xDest) Then Call lstErr_Clear(lstErr, lstTable, xDest & " :existe déjà "): Exit Sub
Set msFile = msFileSystem.GetFile(xSource)
If msFile.Attributes And 1 Then msFile.Attributes = msFile.Attributes - 1

 msFileSystem.MoveFile xSource, xDest

Set msFile = msFileSystem.GetFile(xDest)
msFile.Attributes = msFile.Attributes + 1

If Trim(arrDoc(0).K1) = "" Then arrDoc_SN
arrDoc_DB

ssTab1_Clear
SSTab.Tab = 0
Call lstErr_Clear(lstErr, lstTable, "Création du document : " & arrDoc(0).K1)
Exit Sub
'---------------------------------------------------------
Error_Msg:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "frmElpDoc_cmdOK : " & X)
End Sub


Private Sub cmdPrint_Click()
mnuSelect_Print_Click
End Sub

Private Sub cmdPrint_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdPrint
End Sub

Private Sub cmdPrior_Click()
Dim I As Integer, Nb As Integer
lstSelectFilter.RemoveItem lstSelectFilter.ListCount - 1
cmdPrior.Visible = IIf(lstSelectFilter.ListCount > 0, True, False)

Nb = 0
arrSelectSN_Min = 0

For I = 0 To arrSelectSN_Nb - 1
    If arrSelectSNFilter(I) < lstSelectFilter.ListCount Then Nb = Nb + 1
    If arrSelectSNFilter(I) < lstSelectFilter.ListCount - 1 Then arrSelectSN_Min = arrSelectSN_Min + 1

Next I
arrSelectSN_Nb = Nb


lstSelect_Display

End Sub

Private Sub cmdPrior_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdPrior
End Sub


Private Sub fgDoc_Click()
lstErr.Clear
mnu_Reset

fgDoc_K = fgDoc.Row * fgDoc.Cols
If fgDoc.Row > 1 Then arrDoc_Index = Val(fgDoc.TextArray(2 + fgDoc_K))
If arrDoc_Index > 1 Then
    mnufgDoc_Supprimer = blnInput
    If IsNull(lnkPlan(arrDoc_Index).Memo) Then
        mnufgDoc_Modifier = blnInput
    Else
        mnufgDoc_Modifier = False
    End If
''''    Me.PopupMenu mnufgDoc, vbPopupMenuRightButton
End If

If lstSelect.ListIndex >= 0 And lstSelect.ListIndex < lstSelect.ListCount Then
    mnuDoc_Display = DocAut.Consulter
    mnuDoc_Print = DocAut.Consulter
    
    mnuDossier_Print = DocAut.Consulter 'Saisir
    mnuElpDoc_Print = DocAut.Saisir
End If
PopupMenu mnu

End Sub

Private Sub filDoc_Click()
recElpDoc_Init recDoc
recDoc.Memo = filDoc.FileName
arrDoc_AddItem 0
fgDoc_Display
lstPlan.ListIndex = 1
lstPlan_Click

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case Is = 13: KeyCode = 0:  SendKeys "{TAB}" 'timerControl.Enabled = True '
    Case Is = 27: cmdContext_Quit
    Case Is = 44: KeyCode = 0: frmElpPrt.prtScreen
End Select

End Sub

Private Sub Form_Load()
Set XForm = Me
Call MeInit(arrTagNb)
tableElpTable_Open
tableElpDoc_Open

lstPlan_Load
fgDoc_FormatString = fgDoc.FormatString
fgDoc_BackColorFixed = fgDoc.BackColorFixed
fgDoc_BackColor = fgDoc.BackColor

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub Form_Unload(Cancel As Integer)
'tableElpTable_Close
tableElpDoc_Close
MDB_Local

End Sub

Private Sub fraDétail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub fraSelect_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset

End Sub


Private Sub lstPlan_Click()
Dim X As String
lstErr.Clear
txtMemo.Visible = False
filDoc.Visible = False
lstTable.Visible = False

txtMemo.Enabled = blnInput
'fgDoc.Enabled = blnInput
lstTable.Enabled = blnInput

If blnAddNew Then
    If Not blnAddNew_Document Then
        Call lstErr_Clear(lstErr, lstPlan, "Choisir le document")
        lstPlan.ListIndex = 0
    Else
        If Not blnAddNew_Service Then
            Call lstErr_Clear(lstErr, lstPlan, "Choisir le service")
            lstPlan.ListIndex = 1
        Else
            cmdOk.Visible = True
        End If
    End If
End If

Select Case lstPlan.ListIndex
    Case Is = 0
        If blnAddNew And Not blnAddNew_Document Then
            filDoc.Visible = True
            filDoc.Pattern = "*.*"
        Else
            Call lstErr_Clear(lstErr, lstPlan, "? le nom du document ne peut pas être modifié ")
        End If
     Case Is = 1
        If blnAddNew And Not blnAddNew_Service Then
            lstTable.Visible = True
            X = arrPlan(lstPlan.ListIndex).Memo
            lstTable_Load X
            If lstTable.Enabled Then lstTable.SetFocus
        Else
            Call lstErr_Clear(lstErr, lstPlan, "? le nom du service ne peut pas être modifié ")
        End If
   
    Case Else
        currentMethod = constAddNew
        If IsNull(arrPlan(lstPlan.ListIndex).Memo) Then
            txtMemo.Visible = True
            txtMemo.Text = ""
            If txtMemo.Enabled Then txtMemo.SetFocus
        Else
            lstTable.Visible = True
            X = arrPlan(lstPlan.ListIndex).Memo
            lstTable_Load X
            If lstTable.Enabled Then lstTable.SetFocus
        End If
End Select

End Sub

Private Sub lstPlan_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set lstPlan
End Sub


Private Sub lstSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
mnu_Reset
mnuDossier_Addnew = DocAut.Saisir
If lstSelect.ListCount > 0 Then mnuSelect_Print = DocAut.Saisir

If lstSelect.ListIndex >= 0 And lstSelect.ListIndex < lstSelect.ListCount Then
    mnuDoc_Display = DocAut.Consulter
    mnuDoc_Print = DocAut.Consulter
    
    mnuDossier_Delete = DocAut.Saisir
    mnuDossier_Update = DocAut.Saisir
    mnuDossier_Display = DocAut.Consulter 'Saisir
    mnuDossier_Print = DocAut.Consulter 'Saisir
    mnuElpDoc_Print = DocAut.Saisir
End If
PopupMenu mnu

End Sub

Private Sub lstSelect_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set lstSelect
lstI = lstSelect.TopIndex + Fix(Y / 195)
If lstI >= 0 And lstI < lstSelect.ListCount Then
    lstSelect.ListIndex = lstI
    lstSelect.ToolTipText = lstSelect.Text
    If Y < 195 And lstSelect.TopIndex > 0 Then lstSelect.TopIndex = lstSelect.TopIndex - 1
    If Y > lstSelect.Height - 195 Then lstSelect.TopIndex = lstSelect.TopIndex + 1
End If
End Sub


Private Sub lstSelectPlan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim strX As String
lstErr.Clear

lstSelectTable.Clear
If lstSelectPlan.ListIndex >= 0 And lstSelectPlan.ListIndex < lstSelectPlan.ListCount Then
    If Not IsNull(arrSelectPlan(lstSelectPlan.ListIndex).Memo) Then
        strX = arrSelectPlan(lstSelectPlan.ListIndex).Memo
        lstSelectTable_Load strX
        If lstSelectTable.Enabled Then lstSelectTable.SetFocus
    End If
End If

End Sub


Private Sub lstSelectplan_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set lstSelectPlan
End Sub

Private Sub lstSelectTable_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lstSelectPlan.ListIndex >= 0 And lstSelectPlan.ListIndex < lstSelectPlan.ListCount Then
    If lstSelectTable.ListIndex >= 0 And lstSelectTable.ListIndex < lstSelectTable.ListCount Then
        lstSelectFilter.AddItem arrSelectPlan(lstSelectPlan.ListIndex).K2 & Chr$(9) & " : " & lstSelectTable
        If lstSelectFilter.ListCount = 1 Then
            arrSelectSN_Load arrSelectPlan(lstSelectPlan.ListIndex).K2, arrSelectTable(lstSelectTable.ListIndex).K1
        Else
            arrSelectSN_Filter arrSelectPlan(lstSelectPlan.ListIndex).K2, arrSelectTable(lstSelectTable.ListIndex).K1
        End If
        cmdPrior.Visible = True
        lstSelect_Display
    End If
End If

End Sub


Private Sub lstSelectTable_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set lstSelectTable
End Sub


Private Sub lstTable_Click()
lstErr.Clear
If lstTable_Equal >= 0 Then
    Call lstErr_Clear(lstErr, lstTable, "? ligne déjà sélectionnée")
Else
    recElpDoc_Init recDoc
    recDoc.K2 = arrTable(lstTable.ListIndex).K1
    arrDoc_AddItem 1
    fgDoc_Display
End If
End Sub

Private Sub lstTable_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set lstTable
End Sub


Private Sub mnuDoc_Display_Click()
Dim X As String, xFileName As String, I As Integer

fgDoc_Load arrSelectSN(arrSelectSN_Min + lstSelect.ListIndex)
xFileName = Trim(paramServer(arrDoc(0).Memo))
X = Trim(paramServer(lnkTable(1).Memo)) & "\" & xFileName
I = Len(xFileName)
Select Case UCase$(mId$(xFileName, I - 3, 4))
    Case ".DOC": frmElpPrt.WinWord X
    Case ".XLS": frmElpPrt.Excel X
    Case ".TXT", ".DAT": frmElpPrt.WordPad X
    Case ".BMP": frmElpPrt.MsPaint X
End Select

End Sub

Private Sub mnuDoc_Print_Click()
Dim X As String, xFileName As String, I As Integer

fgDoc_Load arrSelectSN(arrSelectSN_Min + lstSelect.ListIndex)
''xFileName = Trim(CStr(arrDoc(0).Memo))
''X = Trim(CStr(lnkTable(1).Memo)) & "\" & xFileName
xFileName = Trim(paramServer(arrDoc(0).Memo))
X = Trim(paramServer(lnkTable(1).Memo)) & "\" & xFileName

I = Len(xFileName)
Select Case UCase$(mId$(xFileName, I - 3, 4))
    Case ".DOC": frmElpPrt.WinWord_Print X, 2000
    Case ".XLS": frmElpPrt.Excel_Print X, 2000
    Case ".TXT", ".DAT": frmElpPrt.WordPad_Print X, 2000
    Case ".BMP": frmElpPrt.MsPaint_Print X, 2000
End Select

End Sub


Private Sub mnuDossier_Addnew_Click()
ReDim arrDoc(0)
ReDim lnkPlan(0)
ReDim lnkTable(0)

lstTable.Clear
fgDoc.Rows = 1

currentMethod = "AddNew"
blnInput = DocAut.Saisir
blnAddNew = DocAut.Saisir
SSTab.Tab = 1
SSTab.Caption = "Création d'un dossier"
lstPlan.ListIndex = 0: mLstPlan_ListIndex = 0
lstPlan_Click

End Sub


Private Sub mnuDossier_Delete_Click()
Dim X As String
If lstSelect.ListIndex < lstSelect.ListCount _
And DocAut.Saisir Then
    lstErr.Clear
    SSTab.Tab = 1
    SSTab.Caption = "Suppression d'un dossier"
    fgDoc_Load arrSelectSN(arrSelectSN_Min + lstSelect.ListIndex)
    lstPlan.ListIndex = 2: mLstPlan_ListIndex = 0
    X = MsgBox("Voulez-vous réellement supprimer ce dossier?", vbYesNo + vbQuestion + vbDefaultButton2, "Documentation")
    If X = vbYes Then arrDoc_Delete: form_Clear
    SSTab.Tab = 0
End If

End Sub

Private Sub mnuDossier_Display_Click()
If lstSelect.ListIndex < lstSelect.ListCount Then
    blnInput = False
    lstErr.Clear
    SSTab.Tab = 1
    SSTab.Caption = "Affichage d'un dossier"
    fgDoc_Load arrSelectSN(arrSelectSN_Min + lstSelect.ListIndex)
    lstPlan.ListIndex = 2: mLstPlan_ListIndex = 0
End If

End Sub

Private Sub mnuDossier_Print_Click()

If lstSelect.ListIndex >= 0 And lstSelect.ListIndex < lstSelect.ListCount Then
    cmdPrint_Load
    prtElpDocX "Dossier"
End If

End Sub

Private Sub mnuDossier_Update_Click()
If lstSelect.ListIndex < lstSelect.ListCount Then
    blnInput = DocAut.Saisir
    lstErr.Clear
    SSTab.Tab = 1
    SSTab.Caption = "Modification d'un dossier"
    fgDoc_Load arrSelectSN(arrSelectSN_Min + lstSelect.ListIndex)
    lstPlan.ListIndex = 2: mLstPlan_ListIndex = 0
    txtMemo.Visible = False
End If
End Sub


Private Sub mnuElpDoc_Print_Click()
If lstSelect.ListIndex >= 0 And lstSelect.ListIndex < lstSelect.ListCount Then
    cmdPrint_Load
    prtElpDocX "ElpDoc"
End If

End Sub

Private Sub mnufgDoc_Modifier_Click()
currentMethod = constUpdate
txtMemo.Visible = True
txtMemo.Text = Trim(CStr(arrDoc(arrDoc_Index).Memo))
If txtMemo.Enabled Then txtMemo.SetFocus
End Sub

Private Sub mnufgDoc_Supprimer_Click()
If arrDoc(arrDoc_Index).Method = constAddNew Then
    arrDoc(arrDoc_Index).Method = constIgnore
Else
    arrDoc_DBDelete
End If
fgDoc_Display
End Sub



Private Sub mnuSelect_Print_Click()
If lstSelect.ListCount > 0 Then
    cmdPrint_Load
    prtElpDocX "SelectSN"
End If

End Sub

Private Sub cmdServiceX_Click()
Dim intReturn As Integer, mId As String * 12, I As Integer
Dim xService_Old As String, xService_New As String
xService_Old = Trim(InputBox("Ancien code", "Changement d'identifiant d'un service"))
If xService_Old = "" Then Exit Sub
xService_New = Trim(InputBox("Nouveau code", "Changement d'identifiant d'un service"))
If xService_New = "" Then Exit Sub
recElpDoc_Init xElpDoc

xElpDoc.Method = "MoveFirst"

Do
    intReturn = tableElpDoc_Read(xElpDoc)
    If intReturn = 0 Then
        If Trim(xElpDoc.K2) = xService_Old Then
            Select Case Trim(xElpDoc.Id)
                Case "Confidential", "Diffusion", "Rédacteur", "Service"
                        xElpDoc.K2 = xService_New
                        xElpDoc.Method = "Update"
                        If Not IsNull(dbElpDoc_Update(xElpDoc)) Then Call lstErr_Clear(lstErr, lstPlan, "? erreur mise à jour ElpDoc")

            End Select
        End If
        
    End If
    
    xElpDoc.Method = "MoveNext"
  
Loop While intReturn = 0

End Sub

Private Sub SSTab_Click(PreviousTab As Integer)
If SSTab.Tab = 0 Then
    cmdAddNew.Visible = DocAut.Saisir
Else
    cmdAddNew.Visible = False
End If

End Sub

Private Sub SSTab_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset
End Sub


Private Sub txtMemo_GotFocus()
txt_GotFocus txtMemo

End Sub

Private Sub txtMemo_LostFocus()
Dim X As String
txt_LostFocus txtMemo

X = Trim(txtMemo)
If X <> "" Then
    Select Case currentMethod
        Case constAddNew
            recElpDoc_Init recDoc
            recDoc.Memo = X
            recDoc.K2 = arrDoc_Equal
            arrDoc_AddItem 0
        Case constUpdate
            arrDoc(arrDoc_Index).Memo = X
            arrDoc_DBUpdate
            txtMemo.Visible = False
    End Select
    fgDoc_Display
    txtMemo = ""
    If txtMemo.Visible And txtMemo.Enabled Then txtMemo.SetFocus
End If

End Sub

Public Sub lstTable_Load(Id As String)
Dim intReturn As Integer, mId As String * 12, I As Integer
mId = Id
I = 0
ReDim arrTable(10)

lstTable.Visible = True
lstTable.Clear
recElpTable_Init recElpTable

recElpTable.Method = "Seek>="
recElpTable.Id = Id

Do
    intReturn = tableElpTable_Read(recElpTable)
    If intReturn = 0 Then
        If mId <> recElpTable.Id Then
            intReturn = -1
        Else
            lstTable.AddItem recElpTable.Name
            If I = UBound(arrTable) Then ReDim Preserve arrTable(I + 1)
            arrTable(I) = recElpTable
            I = I + 1
        End If
        
    End If
    
    recElpTable.Method = "MoveNext"
  
Loop While intReturn = 0


End Sub
Public Sub lstSelectTable_Load(Id As String)
Dim intReturn As Integer, mId As String * 12, I As Integer
mId = Id
I = 0
ReDim arrSelectTable(10)

recElpTable_Init recElpTable

recElpTable.Method = "Seek>="
recElpTable.Id = Id

Do
    intReturn = tableElpTable_Read(recElpTable)
    If intReturn = 0 Then
        If mId <> recElpTable.Id Then
            intReturn = -1
        Else
            lstSelectTable.AddItem recElpTable.Name
            If I = UBound(arrSelectTable) Then ReDim Preserve arrSelectTable(I + 1)
            arrSelectTable(I) = recElpTable
            I = I + 1
        End If
        
    End If
    
    recElpTable.Method = "MoveNext"
  
Loop While intReturn = 0


End Sub

Public Sub fgDoc_Display()
Dim I As Integer, K As Integer

fgDoc.Visible = True
fgDoc.Clear

fgDoc.Rows = 1
fgDoc.FormatString = fgDoc_FormatString
fgDoc.Enabled = True
For I = 0 To UBound(arrDoc) - 1
    If arrDoc(I).Method <> constDelete _
    And arrDoc(I).Method <> constIgnore Then
        fgDoc.Rows = fgDoc.Rows + 1
        fgDoc.Row = fgDoc.Rows - 1
        K = (fgDoc.Row) * fgDoc.Cols
        fgDoc.TextArray(0 + K) = lnkPlan(I).Name
        fgDoc.TextArray(2 + K) = Format$(I, "####0")
        fgDoc.TextArray(3 + K) = lnkPlan(I).K1
        If Trim(lnkTable(I).Id) <> "" Then
            fgDoc.TextArray(1 + K) = lnkTable(I).Name
            fgDoc.TextArray(4 + K) = lnkTable(I).Name
        Else
             fgDoc.TextArray(1 + K) = arrDoc(I).Memo
            fgDoc.TextArray(4 + K) = arrDoc(I).K2
       End If
  End If
Next I
If fgDoc.Rows > 1 Then fgDoc_Sort

End Sub
Public Sub fgDoc_Sort()
fgDoc.Row = 1
fgDoc.RowSel = fgDoc.Rows - 1

fgDoc.Col = 3
fgDoc.ColSel = fgDoc.Cols - 1
fgDoc.Sort = 1

End Sub


Public Sub lstPlan_Load()
Dim intReturn As Integer, mId As String * 12, I As Integer, ISelect As Integer
mId = "Doc_Plan"
I = 0
ReDim arrPlan(10)
lstPlan.Clear

ISelect = 0
ReDim arrSelectPlan(10)
lstSelectPlan.Clear

recElpTable_Init recElpTable

recElpTable.Method = "Seek>="
recElpTable.Id = mId

Do
    intReturn = tableElpTable_Read(recElpTable)
    If intReturn = 0 Then
        If mId <> recElpTable.Id Then
            intReturn = -1
        Else
            lstPlan.AddItem recElpTable.Name
            If I = UBound(arrPlan) Then ReDim Preserve arrPlan(I + 1)
            arrPlan(I) = recElpTable
            I = I + 1
            
            If Not IsNull(recElpTable.Memo) Then
                lstSelectPlan.AddItem recElpTable.Name
                If ISelect = UBound(arrSelectPlan) Then ReDim Preserve arrPlan(ISelect + 1)
                arrSelectPlan(ISelect) = recElpTable
                ISelect = ISelect + 1
            End If
        End If
        
    End If
    
    recElpTable.Method = "MoveNext"
  
Loop While intReturn = 0

End Sub


Public Sub arrSelectSN_Load(xId As String, xK2 As String)
Dim intReturn As Integer, mId As String * 12, mK2 As String * 12
ReDim arrSelectSN(0)
ReDim arrSelectSNFilter(0)

mId = xId: mK2 = xK2
arrSelectSN_Nb = 0
arrSelectSN_Min = 0

recElpDoc_Init recElpDoc

recElpDoc.Method = "Seek>="
recElpDoc.Id = mId

Do
    intReturn = tableElpDoc_Read(recElpDoc)
    If intReturn = 0 Then
        If mId <> recElpDoc.Id Then
            intReturn = -1
        Else
            If recElpDoc.K2 = mK2 Then
                If Confidential_Control Then
 '                   lstSelect.AddItem recElpDoc.Memo
                    If arrSelectSN_Nb = UBound(arrSelectSN) Then
                        ReDim Preserve arrSelectSN(arrSelectSN_Nb + 1)
                        ReDim Preserve arrSelectSNFilter(arrSelectSN_Nb + 1)
                    End If
                    arrSelectSN(arrSelectSN_Nb) = recElpDoc.K1
                    arrSelectSNFilter(arrSelectSN_Nb) = 0
                    arrSelectSN_Nb = arrSelectSN_Nb + 1
                End If
            End If
        End If
        
    End If
    
    recElpDoc.Method = "Seek>"
  
Loop While intReturn = 0

End Sub
Public Function Confidential_Control() As Boolean
Dim intReturn As Integer, X3 As String

Confidential_Control = True
If DocAut.Saisir Then Exit Function

recElpDoc_Init xElpDoc

xElpDoc.Method = "Seek>="
xElpDoc.Id = "Confidential"
xElpDoc.K1 = recElpDoc.K1

Do
    intReturn = tableElpDoc_Read(xElpDoc)
    If intReturn = 0 Then
        If "Confidential" <> xElpDoc.Id Or xElpDoc.K1 <> recElpDoc.K1 Then
            intReturn = -1
        Else
            Confidential_Control = False
            If xElpDoc.K2 = usrId Then Confidential_Control = True: Exit Function
            X3 = mId$(xElpDoc.K2, 1, 3)
            If X3 >= usrService_Min And X3 <= usrService_Max Then Confidential_Control = True: Exit Function
        End If
        
    End If
    
    xElpDoc.Method = "MoveNext"
  
Loop While intReturn = 0
End Function



Public Sub arrSelectSN_Filter(xId As String, xK2 As String)
Dim intReturn As Integer, mId As String * 12, mK2 As String * 12, I As Integer, Nb As Integer

mId = xId: mK2 = xK2
recElpDoc_Init recElpDoc

recElpDoc.Method = "Seek="
recElpDoc.Id = mId
recElpDoc.K2 = mK2
Nb = arrSelectSN_Nb - 1
arrSelectSN_Min = arrSelectSN_Nb

For I = 0 To Nb
    If arrSelectSNFilter(I) = lstSelectFilter.ListCount - 2 Then
        recElpDoc.K1 = arrSelectSN(I)
        intReturn = tableElpDoc_Read(recElpDoc)
        If intReturn = 0 Then
            If recElpDoc.K2 = mK2 Then
                lstSelect.AddItem recElpDoc.Memo
                 If arrSelectSN_Nb = UBound(arrSelectSN) Then
                     ReDim Preserve arrSelectSN(arrSelectSN_Nb + 1)
                     ReDim Preserve arrSelectSNFilter(arrSelectSN_Nb + 1)
                 End If
                 arrSelectSN(arrSelectSN_Nb) = recElpDoc.K1
                 arrSelectSNFilter(arrSelectSN_Nb) = lstSelectFilter.ListCount - 1
                 arrSelectSN_Nb = arrSelectSN_Nb + 1
            End If
        End If
    End If
Next I

End Sub



Public Sub fgDoc_Load(K1 As String)
Dim intReturn As Integer, mId As String * 12, I As Integer, IPlan As Integer
I = 0
ReDim arrDoc(0), lnkPlan(0), lnkTable(0)
fgDoc.Clear
recElpDoc_Init recElpDoc

For IPlan = 0 To UBound(arrPlan) - 1
    recElpDoc.Method = "Seek>="
    mId = arrPlan(IPlan).K2
    recElpDoc.Id = mId
    recElpDoc.K1 = K1
    recElpDoc.K2 = ""
    Do
        intReturn = tableElpDoc_Read(recElpDoc)
        If intReturn = 0 Then
            If mId <> recElpDoc.Id Or recElpDoc.K1 <> K1 Then
                intReturn = -1
            Else
                If I = UBound(arrDoc) Then
                    ReDim Preserve arrDoc(I + 1)
                    ReDim Preserve lnkPlan(I + 1)
                    ReDim Preserve lnkTable(I + 1)
                End If
                   
                   arrDoc(I) = recElpDoc
                lnkPlan(I) = arrPlan(IPlan)
                If IsNull(arrPlan(IPlan).Memo) Then
                    lnkTable(I).Id = ""
                Else
                    lnkTable(I).Id = CStr(arrPlan(IPlan).Memo)
                    lnkTable(I).K1 = recElpDoc.K2
                    lnkTable(I).K2 = ""
                    lnkTable(I).Method = "Seek="
                    tableElpTable_Read lnkTable(I)
                End If

                I = I + 1
            End If
            
        End If
        
        recElpDoc.Method = "MoveNext"
      
    Loop While intReturn = 0
Next IPlan
fgDoc_Display
End Sub

Public Sub arrDoc_AddItem(kLnk As Integer)
Dim I As Integer
recDoc.Method = currentMethod
recDoc.Id = arrPlan(lstPlan.ListIndex).K2

I = UBound(arrDoc)
ReDim Preserve arrDoc(I + 1)
ReDim Preserve lnkTable(I + 1)
ReDim Preserve lnkPlan(I + 1)

arrDoc(I) = recDoc

lnkPlan(I) = arrPlan(lstPlan.ListIndex)
Select Case kLnk
    Case 1: lnkTable(I) = arrTable(lstTable.ListIndex)
    Case Else: lnkTable(I).Id = ""
End Select

If blnAddNew Then
    If lstPlan.ListIndex = 1 Then
        blnAddNew_Service = True: lstPlan.ListIndex = 2
    Else
        If lstPlan.ListIndex = 0 Then
            blnAddNew_Document = True: lstPlan.ListIndex = 1
        Else
            If mLstPlan_ListIndex <> lstPlan.ListIndex Then
                    mLstPlan_ListIndex = lstPlan.ListIndex
                    If lstPlan.ListIndex < lstPlan.ListCount - 1 Then lstPlan.ListIndex = lstPlan.ListIndex + 1
            End If
        End If
    End If
Else
    arrDoc(I).K1 = arrDoc(0).K1
    arrDoc(I).Method = constAddNew
    If Not IsNull(dbElpDoc_Update(arrDoc(I))) Then
        Call lstErr_Clear(lstErr, lstPlan, "? erreur mise à jour ElpDoc")
    Else
        arrDoc(I).Method = ""
    End If
End If
End Sub

Public Sub arrDoc_SN()
lstErr.Clear
recElpTable_Init recElpTable
recElpTable.Id = "Doc_Param"
recElpTable.K1 = "SN"
recElpTable.K2 = ""
recElpTable.Method = "Seek="
If Not IsNull(dbElpTable_ReadE(recElpTable)) Then
    Call lstErr_Clear(lstErr, lstPlan, "? erreur lecture SN")
Else
    recElpTable.SNP = recElpTable.SNP + 1
    recElpTable.Method = "Update"
    If Not IsNull(dbElpTable_Update(recElpTable)) Then
        Call lstErr_Clear(lstErr, lstPlan, "? erreur mise à jour SN")
    Else
        arrDoc(0).K1 = Format$(recElpTable.SNP, "000000000000")
    End If
End If

End Sub

Public Sub arrDoc_DB()
Dim I As Integer


For I = 0 To UBound(arrDoc) - 1
    If Trim(arrDoc(I).Method) <> "" _
    And arrDoc(I).Method <> constIgnore Then
        arrDoc(I).K1 = arrDoc(0).K1
        If Not IsNull(dbElpDoc_Update(arrDoc(I))) Then
            Call lstErr_Clear(lstErr, lstPlan, "? erreur mise à jour ElpDoc")
            Exit For
        End If
    End If
Next I

End Sub

Public Function lstTable_Equal() As Integer
For lstTable_Equal = 0 To UBound(arrDoc) - 1
    If arrDoc(lstTable_Equal).Id = arrPlan(lstPlan.ListIndex).K2 Then
        If arrDoc(lstTable_Equal).Method <> constIgnore Then
            If arrDoc(lstTable_Equal).K2 = arrTable(lstTable.ListIndex).K1 Then Exit Function
        End If
    End If
Next lstTable_Equal
lstTable_Equal = -1
End Function

Public Function arrDoc_Equal() As String
Dim I As Integer, X12 As String * 12
X12 = "000000000000"
For I = 0 To UBound(arrDoc) - 1
    If arrDoc(I).Id = arrPlan(lstPlan.ListIndex).K2 Then
        If arrDoc(I).K2 > X12 Then X12 = arrDoc(I).K2
        End If
Next I
arrDoc_Equal = Format$(Val(X12) + 1, "000000000000")
End Function


Public Sub arrDoc_Delete()
Dim xSrc As String, X12 As String * 12, xFileName As String

On Error GoTo Error_Msg
X12 = arrDoc(0).K1
For arrDoc_Index = 0 To UBound(arrDoc) - 1
    arrDoc_DBDelete
Next arrDoc_Index

''X = Trim(CStr(arrDoc(0).Memo))
''xSrc = Trim(CStr(lnkTable(1).Memo)) & "\" & X
xFileName = Trim(paramServer(arrDoc(0).Memo))
xSrc = Trim(paramServer(lnkTable(1).Memo)) & "\" & xFileName


If Dir(xSrc, vbReadOnly + vbHidden) = "" Then Call lstErr_Clear(lstErr, lstTable, xSrc & " :n'existe pas "): Exit Sub

Name xSrc As paramElpDoc_Archive & "\" & X12 & "_" & xFileName
Exit Sub
'---------------------------------------------------------
Error_Msg:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Document_arrDoc_Delete : " & xFileName)
End Sub

Public Sub arrDoc_DBUpdate()
recElpDoc = arrDoc(arrDoc_Index)
recElpDoc.Method = "Seek="
If Not IsNull(dbElpDoc_Read(recElpDoc)) Then
    Call lstErr_Clear(lstErr, lstPlan, "? erreur lecture ElpDoc")
Else
    arrDoc(arrDoc_Index).Method = constUpdate
    If Not IsNull(dbElpDoc_Update(arrDoc(arrDoc_Index))) Then
        Call lstErr_Clear(lstErr, lstPlan, "? erreur mise à jour ElpDoc")
    Else
        arrDoc(arrDoc_Index).Method = ""
    End If
End If

End Sub

Public Sub arrDoc_DBDelete()
recElpDoc = arrDoc(arrDoc_Index)
recElpDoc.Method = "Seek="
If Not IsNull(dbElpDoc_Read(recElpDoc)) Then
    Call lstErr_Clear(lstErr, lstPlan, "? erreur lecture ElpDoc")
Else
    arrDoc(arrDoc_Index).Method = constDelete
    If Not IsNull(dbElpDoc_Update(arrDoc(arrDoc_Index))) Then
        Call lstErr_Clear(lstErr, lstPlan, "? erreur mise à jour ElpDoc")
    Else
        arrDoc(arrDoc_Index).Method = constIgnore
    End If
End If

End Sub


Public Sub cmdPrint_Load()
Dim I As Integer

lstErr.Clear
fgDoc_Load arrSelectSN(arrSelectSN_Min + lstSelect.ListIndex)

ReDim prtElpDoc.arrPlan(UBound(arrPlan))
For I = 0 To UBound(arrPlan)
    prtElpDoc.arrPlan(I) = arrPlan(I)
Next I

ReDim prtElpDoc.lnkPlan(UBound(lnkPlan))
For I = 0 To UBound(lnkPlan)
    prtElpDoc.lnkPlan(I) = lnkPlan(I)
Next I

ReDim prtElpDoc.lnkTable(UBound(lnkTable))
For I = 0 To UBound(lnkTable)
    prtElpDoc.lnkTable(I) = lnkTable(I)
Next I

ReDim prtElpDoc.arrDoc(UBound(arrDoc))
For I = 0 To UBound(arrDoc)
    prtElpDoc.arrDoc(I) = arrDoc(I)
Next I
'X = arrSelect(arrSelectSN_Min + lstSelect.ListIndex).K1

If lstSelectFilter.ListCount <= 0 Then
    ReDim prtElpDoc.arrSelectFilter(0)
Else
    ReDim prtElpDoc.arrSelectFilter(lstSelectFilter.ListCount)
    For I = 0 To lstSelectFilter.ListCount - 1
        lstSelectFilter.ListIndex = I
        prtElpDoc.arrSelectFilter(I) = lstSelectFilter.Text
    Next I
End If

If lstSelect.ListCount <= 0 Then
    ReDim prtElpDoc.arrSelectSN(0)
Else
    ReDim prtElpDoc.arrSelectSN(lstSelect.ListCount)
    For I = 0 To lstSelect.ListCount - 1
        prtElpDoc.arrSelectSN(I) = arrSelectSN(arrSelectSN_Min + I)
    Next I
End If

End Sub

Public Sub ssTab1_Clear()
blnMsgBox_Quit = False
blnInput = False
blnAddNew = False
blnAddNew_Document = False
blnAddNew_Service = False
currentMethod = ""

fgDoc.Clear
lstTable.Clear
SSTab.Caption = "Recherche"
cmdOk.Visible = False
lstTable.Visible = False
txtMemo.Visible = False
txtMemo.Top = lstTable.Top
txtMemo.Left = lstTable.Left
txtMemo.Text = ""
filDoc.Visible = False
filDoc.Top = lstTable.Top
filDoc.Left = lstTable.Left
filDoc.Pattern = "*.*"
filDoc.Path = paramElpDoc_Validation

End Sub

Public Sub lstSelect_Display()
Dim I As Integer

lstSelect.Clear
recElpDoc_Init recElpDoc

recElpDoc.Method = "Seek="

For I = 0 To arrSelectSN_Nb - 1
    If arrSelectSNFilter(I) = lstSelectFilter.ListCount - 1 Then
        recElpDoc.Id = "Intitulé"
        recElpDoc.K2 = "000000000001"
        recElpDoc.K1 = arrSelectSN(I)
        If tableElpDoc_Read(recElpDoc) <> 0 Then
            recElpDoc.Id = "Document"
            recElpDoc.K2 = ""
            recElpDoc.K1 = arrSelectSN(I)
            If tableElpDoc_Read(recElpDoc) <> 0 Then recElpDoc.Memo = "???? err " & arrSelectSN(I)
        End If
        lstSelect.AddItem recElpDoc.Memo
    End If
Next I

End Sub

Public Sub mnu_Reset()
mnuDoc_Display = False
mnuDoc_Print = False

mnuDossier_Addnew = False
mnuDossier_Delete = False
mnuDossier_Update = False
mnuDossier_Display = False
mnuDossier_Print = False
mnuElpDoc_Print = False
mnuSelect_Print = False

mnufgDoc_Supprimer = False
mnufgDoc_Modifier = False

End Sub
