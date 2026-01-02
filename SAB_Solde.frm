VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSAB_Solde 
   AutoRedraw      =   -1  'True
   Caption         =   "SAB_Comptabilité : Interface SOLDES"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9090
   Icon            =   "SAB_Solde.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6210
   ScaleWidth      =   9090
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   450
      Left            =   4350
      TabIndex        =   5
      Top             =   -30
      Width           =   4260
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5445
      Left            =   0
      TabIndex        =   3
      Top             =   495
      Width           =   9090
      _ExtentX        =   16034
      _ExtentY        =   9604
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Interface Soldes"
      TabPicture(0)   =   "SAB_Solde.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraTab0"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "xxxxx"
      TabPicture(1)   =   "SAB_Solde.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraTAb1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Maintenance informatique"
      TabPicture(2)   =   "SAB_Solde.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraInfo"
      Tab(2).ControlCount=   1
      Begin VB.Frame fraInfo 
         Height          =   4905
         Left            =   -74910
         TabIndex        =   11
         Top             =   480
         Width           =   8820
         Begin VB.TextBox txtInfo 
            Height          =   285
            Left            =   2145
            TabIndex        =   17
            Top             =   465
            Width           =   4020
         End
         Begin VB.CommandButton cmdInfo 
            BackColor       =   &H00C0C0FF&
            Caption         =   "Lancer le traitement"
            Height          =   810
            Left            =   6285
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   2070
            Width           =   2400
         End
         Begin VB.TextBox txtInfoExport 
            Height          =   285
            Left            =   2130
            TabIndex        =   13
            Top             =   1620
            Width           =   4020
         End
         Begin VB.TextBox txtInfoImport 
            Height          =   285
            Left            =   2130
            TabIndex        =   12
            Top             =   1110
            Width           =   4020
         End
         Begin VB.Label lblInfo 
            Caption         =   "***"
            Height          =   255
            Left            =   210
            TabIndex        =   18
            Top             =   510
            Width           =   1155
         End
         Begin VB.Label lblInfoExport 
            Caption         =   "Fichier export"
            Height          =   255
            Left            =   195
            TabIndex        =   15
            Top             =   1710
            Width           =   1335
         End
         Begin VB.Label lblInfoImport 
            Caption         =   "Fichier import"
            Height          =   225
            Left            =   165
            TabIndex        =   14
            Top             =   1110
            Width           =   1395
         End
      End
      Begin VB.Frame fraTAb1 
         Height          =   4905
         Left            =   -74940
         TabIndex        =   10
         Top             =   390
         Width           =   8940
      End
      Begin VB.Frame fraTab0 
         Height          =   4875
         Left            =   135
         TabIndex        =   4
         Top             =   360
         Width           =   8850
         Begin VB.CommandButton cmdOptions 
            BackColor       =   &H00F0FFFF&
            Caption         =   "Options"
            Height          =   645
            Left            =   195
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   390
            Width           =   810
         End
         Begin VB.Frame fraContextOptions 
            BackColor       =   &H00F0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4350
            Left            =   5010
            TabIndex        =   19
            Top             =   120
            Width           =   3795
            Begin VB.FileListBox filDoc 
               BackColor       =   &H00F0FFFF&
               ForeColor       =   &H00008000&
               Height          =   3405
               Left            =   165
               TabIndex        =   22
               Top             =   795
               Width           =   3500
            End
            Begin VB.Label lblContextOptions 
               BackColor       =   &H00F0FFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Label1"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   615
               Left            =   180
               TabIndex        =   20
               Top             =   120
               Width           =   3480
            End
         End
         Begin VB.Frame fraSelect 
            Height          =   870
            Left            =   1080
            TabIndex        =   6
            Top             =   225
            Width           =   6540
            Begin VB.CommandButton cmdSelect 
               BackColor       =   &H00C0FFC0&
               Caption         =   "lecture fichier YMOUVEA0"
               Height          =   645
               Left            =   5085
               Style           =   1  'Graphical
               TabIndex        =   7
               Top             =   180
               Width           =   1065
            End
            Begin VB.TextBox txtSelect 
               Height          =   285
               Left            =   75
               TabIndex        =   0
               Top             =   510
               Width           =   4815
            End
            Begin VB.Label lblSelect 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Fichier à importer"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   135
               TabIndex        =   8
               Top             =   180
               Width           =   4755
            End
         End
         Begin MSFlexGridLib.MSFlexGrid fgSelect 
            Height          =   3585
            Left            =   135
            TabIndex        =   9
            Top             =   1200
            Width           =   8595
            _ExtentX        =   15161
            _ExtentY        =   6324
            _Version        =   393216
            Rows            =   1
            Cols            =   10
            FixedCols       =   0
            RowHeightMin    =   200
            BackColor       =   14737632
            ForeColor       =   12582912
            ForeColorFixed  =   -2147483641
            BackColorSel    =   12648384
            BackColorBkg    =   14737632
            AllowBigSelection=   0   'False
            TextStyle       =   4
            TextStyleFixed  =   4
            FocusRect       =   2
            HighLight       =   0
            GridLinesFixed  =   1
            AllowUserResizing=   3
            FormatString    =   $"SAB_Solde.frx":035E
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
      TabIndex        =   2
      Top             =   0
      Width           =   1200
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   500
      Left            =   8600
      Picture         =   "SAB_Solde.frx":03E8
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   500
   End
   Begin VB.Menu mnuContext 
      Caption         =   "mnuContext"
      Visible         =   0   'False
      Begin VB.Menu mnuContextOptions 
         Caption         =   "Options"
      End
      Begin VB.Menu mnuContextX1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuContextAbandonner 
         Caption         =   "Abandonner"
      End
      Begin VB.Menu mnuContextQuitter 
         Caption         =   "Quitter"
      End
   End
   Begin VB.Menu mnuPrint0 
      Caption         =   "mnuPrint0"
      Visible         =   0   'False
      Begin VB.Menu mnuSelect_Print 
         Caption         =   "Imprimer le journal"
      End
      Begin VB.Menu mnuSelect_Print_Recap 
         Caption         =   "Imprimer les totaux"
      End
   End
End
Attribute VB_Name = "frmSAB_Solde"
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
Dim SAB_Solde_Aut As typeAuthorization
Dim blnAuto As Boolean, blnError As Boolean

Dim fgSelect_FormatString As String, fgSelect_K As Integer
Dim fgSelect_RowDisplay As Integer, fgSelect_RowClick As Integer, fgSelect_ColClick As Integer
Dim fgSelect_ColorClick As Long, fgSelect_ColorDisplay As Long
Dim fgSelect_Sort1 As Integer, fgSelect_Sort2 As Integer
Dim fgSelect_SortAD As Integer, fgSelect_Sort1_Old As Integer
Dim fgSelect_arrIndex As Integer
Dim blnfgSelect_DisplayLine As Boolean


Dim paramSAB_Solde_Import As String
Dim paramSAB_Solde_Archive As String
Dim paramSAB_Solde_Folder As String, paramSAB_Solde_Name As String, paramSAB_Solde_Extension As String

Dim meYSOLDE0 As typeYSOLDE0, xYSOLDE0 As typeYSOLDE0
Dim xYCOMPTE0 As typeYCOMPTE0

Dim xMvtP0 As typeMvtP0
Public Sub cmdAuto()
cmdSelect_Click
'If Not blnError Then cmdReuters_Ok_Click
'Shell_MsgBox "# Reuters => SAB : mise à jour des taux & cours terminée # " & Time, vbInformation, Me.Caption, True

Unload Me
End Sub

Private Sub fgSelect_Display()
Dim Nb As Long
SSTab1.Tab = 0

fgSelect_Reset

fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
fgSelect.Visible = False

xMvtP0.Id = constYCOMPTE0
xMvtP0.Method = "Seek>="
intReturn = tableMvtP0_Read(xMvtP0)

Do

    fgSelect_DisplayLine
    
    xMvtP0.Method = "Seek>"
    intReturn = tableMvtP0_Read(xMvtP0)
    If Trim(mId$(xMvtP0.Id, 1, 10)) <> constYCOMPTE0 Then intReturn = -1
    Nb = Nb + 1
    If Nb Mod 500 = 0 Then Call lstErr_ChangeLastItem(Me.lstErr, Me.cmdContext, "affichage : " & Nb)

Loop Until intReturn <> 0

fgSelect.Visible = True

fgSelect_Sort_Options

End Sub


Public Sub fgSelect_DisplayLine()
Dim X As String
On Error Resume Next
    MsgTxt = Space$(34) & xMvtP0.Text
    MsgTxtIndex = 0
    If Trim(mId$(xMvtP0.Id, 1, 10)) = constYCOMPTE0 Then
    

            srvYCOMPTE0_GetBuffer xYCOMPTE0
            
            fgSelect.Rows = fgSelect.Rows + 1
            fgSelect.Row = fgSelect.Rows - 1
            fgSelect.Col = 0: fgSelect.Text = xYCOMPTE0.COMPTEOBL
            fgSelect.Col = 1: fgSelect.Text = xYCOMPTE0.COMPTEDEV
            fgSelect.Col = 2: fgSelect.Text = xYCOMPTE0.COMPTECOM
            fgSelect.Col = 3: 'fgSelect.Text = xYCOMPTE0.COMPTEcen
            fgSelect.Col = fgSelect_arrIndex: fgSelect.Text = Trim(xMvtP0.Id)
     End If
End Sub


Public Sub fgSelect_Sort()
fgSelect.Visible = False
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
fgSelect.Visible = True
End Sub
Public Sub fgSelect_SortX(lK As Integer)
Dim I As Integer, X As String
Dim mK As Integer
mK = lK
For I = 1 To fgSelect.Rows - 1
    fgSelect.Row = I
    fgSelect.Col = fgSelect_arrIndex
    Select Case lK
        Case 5:
            fgSelect.Col = 0: X = fgSelect.Text
            fgSelect.Col = 2: X = X & fgSelect.Text
            fgSelect.Col = 5: X = X & fgSelect.Text
            fgSelect.Col = 3: X = X & fgSelect.Text
            fgSelect.Col = 4: X = X & fgSelect.Text
         Case 6
            fgSelect.Col = 6: X = fgSelect.Text
            fgSelect.Col = 5: X = X & fgSelect.Text
            fgSelect.Col = 3: X = X & fgSelect.Text
            fgSelect.Col = 4: X = X & fgSelect.Text
    End Select
    fgSelect.Col = fgSelect_arrIndex - 1
    fgSelect.Text = X
Next I


fgSelect_Sort1 = fgSelect_arrIndex - 1: fgSelect_Sort2 = fgSelect_arrIndex - 1
fgSelect_Sort
fgSelect_Sort1 = mK: fgSelect_Sort2 = mK
End Sub



'---------------------------------------------------------
Public Sub arrTag_Set(ByVal B As Boolean)
'---------------------------------------------------------
For I = 0 To arrTagNb: arrTag(I) = B: Next I
End Sub


Public Sub Msg_Rcv(Msg As String)
'---------------------------------------------------------
mWindowState = Me.WindowState
If Me.WindowState <> frmUsr_Windowstate Then Me.WindowState = frmUsr_Windowstate

Call BiaPgmAut_Init("SAB_Solde", SAB_Solde_Aut)

'blnSetfocus = True
Form_Init

Select Case UCase$(Trim(mId$(Msg, 1, 12)))
    Case "@AUTO_TAU":     blnAuto = True: cmdAuto
    Case Else: blnAuto = False
End Select

End Sub


Public Sub Form_Init()
Me.Enabled = False
Call lstErr_Clear(lstErr, cmdPrint, "Initialisation ")
DoEvents
lstErr.Visible = True
If Not IsNull(param_Init) Then
    MsgBox "paramétrage inconsistent", vbCritical, "frmSAB_YSOLDE0.param_init"
    fraTab0.Enabled = False
End If

    blnControl = False
    fgSelect_FormatString = fgSelect.FormatString
    fgSelect.Enabled = True
    fraTAb1.Visible = SAB_Solde_Aut.Xspécial
    cmdReset
Me.Enabled = True

End Sub


'---------------------------------------------------------
Public Sub cmdReset()
'---------------------------------------------------------

blnControl = False
blnError = False
usrColor_Set
cmdContext.Caption = constcmdRechercher: blnMsgBox_Quit = False
arrTag_Set False
currentAction = ""
SSTab1.Tab = 0
lblContextOptions.Caption = "Sélectionner les options" & Asc10_13 & "'Esc' ou cliquer ICI pour quitter"
lblContextOptions.ForeColor = warnUsrColor
mnuContextOptions_Click

txtSelect = paramSAB_Solde_Name & "." & paramSAB_Solde_Extension

txtSelect.Enabled = SAB_Solde_Aut.Xspécial
lblSelect.ForeColor = vbYellow 'warnUsrColor

fraInfo.Enabled = SAB_Solde_Aut.Xspécial
txtInfo = paramTemp_Folder & "FTP\YSOLDE0.txt"

blnControl = True

End Sub



Public Function param_Init()
Dim K As Integer, K1 As Integer, X As String

Dim V
param_Init = Null
On Error GoTo Table_Error
Call lstErr_Clear(Me.lstErr, Me.cmdContext, "BIA.mdb : table : " & recElpTable.Id)


recElpTable.Method = "Seek="
recElpTable.Id = "SAB_Compta"

recElpTable.K1 = "Archive"
recElpTable.K2 = paramIBM_Library_SAB

V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramSAB_Solde_Archive = paramServer(recElpTable.Memo)
If SAB_Solde_Aut.Xspécial Then Call lstErr_AddItem(Me.lstErr, Me.cmdContext, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

paramSAB_Solde_Import = paramSAB_Solde_Archive & fileName_AMJCPT(constYSOLDE0, 0)
Call fileName_Split(paramSAB_Solde_Import, paramSAB_Solde_Folder, paramSAB_Solde_Name, paramSAB_Solde_Extension)

Exit Function

Table_Error:
param_Init = V
Shell_MsgBox "frmSAB_YSOLDE0.param_init#  " & recElpTable.Id & " : " & recElpTable.K1 & " : " & recElpTable.K2 & V, vbCritical, Me.Caption, blnAuto
Exit Function

Memo_Error:
param_Init = "Memo"
Shell_MsgBox "frmSAB_YSOLDE0.param_init#  " & recElpTable.Id & " : " & recElpTable.K1 & " : " & recElpTable.K2 & " manque memo", vbCritical, Me.Caption, blnAuto
Exit Function

End Function
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

Private Sub cmdInfo_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass

Call lstErr_Clear(lstErr, cmdContext, "cmdInfo  : ")

Call lstErr_AddItem(lstErr, cmdContext, "cmdInfo : ")

Me.Enabled = True: Me.MousePointer = 0


End Sub

Private Sub cmdOptions_Click()
If fraContextOptions.Visible Then
    fraContextOptions_Exit
Else
    mnuContextOptions_Click
End If
End Sub

Public Sub fraContextOptions_Exit()
fraContextOptions.Visible = False

End Sub

Private Sub cmdPrint_Click()
Msg = Space$(50)

Select Case SSTab1.Tab
    Case 0:
            If fgSelect.Rows > 1 Then
                fgSelect_Sort_Options
                Me.PopupMenu mnuPrint0, vbPopupMenuLeftButton
           End If
End Select

End Sub

Private Sub lblContextOptions_Click()
fraContextOptions_Exit
End Sub

Private Sub mnuSelect_Print_Click()
'blnTotal = True
cmdPrint_Journal

End Sub

Private Sub cmdSelect_Click()
On Error Resume Next

Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> SAB_Solde_Import"): DoEvents

mdbMvtP0.tableMvtP0_Close
MDB.Execute "delete * from mvtp0"
mdbMvtP0.tableMvtP0_Open

mdbElpKMInfo.tableElpKMInfo_Open


cmdYSOLDE0_Import

cmdYCOMPTE0_Import

Call lstErr_AddItem(lstErr, cmdContext, "! préparation affichage ...... "): DoEvents

fgSelect_Display

Call lstErr_ChangeLastItem(lstErr, cmdContext, "= SAB_Solde_Import"): DoEvents

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub fgSelect_Click()
fgSelect.LeftCol = 0

End Sub

Private Sub fgSelect_LeaveCell()
On Error Resume Next
'fgSelect.CellBackColor = &HE0E0E0
End Sub

Private Sub filDoc_Click()
On Error Resume Next
fraContextOptions.Visible = False
txtSelect = filDoc.FileName
cmdSelect_Click
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

If fraContextOptions.Visible Then fraContextOptions_Exit: Exit Sub
If SSTab1.Tab = 0 Then
        Unload Me
    Exit Sub
End If


If currentAction = "" Then
   
Else
    X = MsgBox("Voulez-vous réellement abandonner la mise à jour?", vbYesNo + vbQuestion + vbDefaultButton2, Me.Caption)
    If X = vbYes Then
        currentAction = ""
    Else
        Exit Sub
    End If
End If

End Sub
Private Sub mnuContextOptions_Click()
filDoc.Path = paramSAB_Solde_Archive
filDoc.Pattern = "*.XXX"
filDoc.Pattern = "*" & Trim(constYSOLDE0) & "*." & paramSAB_Solde_Extension
filDoc.Visible = True
fraContextOptions.Visible = True

End Sub

Public Sub cmdContext_Return()
If SSTab1.Tab = 0 Then
    cmdSelect_Click
Else
    SendKeys "{TAB}"
End If
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





Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim wOrigine As String
Dim V

On Error Resume Next
If Y <= fgSelect.RowHeightMin Then
    Select Case fgSelect.Col
        Case 0: fgSelect_Sort1 = 0: fgSelect_Sort2 = 4: fgSelect_Sort
        Case 1:  fgSelect_Sort1 = 1: fgSelect_Sort2 = 4: fgSelect_Sort
        Case 2: fgSelect_Sort1 = 2: fgSelect_Sort2 = 4: fgSelect_Sort
        Case 3: fgSelect_Sort1 = 3: fgSelect_Sort2 = 4: fgSelect_Sort
        Case 4: fgSelect_Sort1 = 4: fgSelect_Sort2 = 4: fgSelect_Sort
        Case 5: fgSelect_Sort1 = 5: fgSelect_Sort2 = 5: fgSelect_SortX fgSelect_Sort1
        Case 6: fgSelect_Sort1 = 6: fgSelect_Sort2 = 6: fgSelect_SortX fgSelect_Sort1
        Case fgSelect_arrIndex:  fgSelect_SortX fgSelect_arrIndex
    End Select
Else
    If fgSelect.Rows > 1 Then
        Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
        fgSelect.Col = fgSelect_arrIndex
        xMvtP0.Id = fgSelect.Text
        xMvtP0.Method = "Seek="
        If tableMvtP0_Read(xMvtP0) = 0 Then
            MsgTxt = Space$(34) & xMvtP0.Text
            MsgTxtIndex = 0
                      

            Shell_MsgBox "fgSelect_MouseDown# " & xMvtP0.Id & " : " & xMvtP0.Err, vbCritical, Me.Caption, False

        End If
    End If
   End If
End Sub

Public Sub fgSelect_Reset()
fgSelect.Clear
fgSelect_Sort1 = 0: fgSelect_Sort2 = 4
fgSelect_Sort1_Old = -1
fgSelect_RowDisplay = 0: fgSelect_RowClick = 0
fgSelect_arrIndex = 9
blnfgSelect_DisplayLine = False
fgSelect_SortAD = 6
fgSelect.LeftCol = 0

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

End Sub




Public Sub cmdYSOLDE0_Import()
Dim wFileName As String
Dim xIn As String, X As String
Dim seq As Long
On Error GoTo Error_Handler

wFileName = paramSAB_Solde_Folder & Trim(txtSelect)

Call lstErr_AddItem(lstErr, cmdContext, constYSOLDE0): DoEvents
recMvtP0_Init xMvtP0
xMvtP0.Method = constAddNew

seq = 0
Open wFileName For Input As #1

Do Until EOF(1)
    seq = seq + 1
    
    If seq Mod 500 = 0 Then Call lstErr_ChangeLastItem(Me.lstErr, Me.cmdContext, constYSOLDE0 & " : " & seq)
    DoEvents
    Line Input #1, xIn
    If xIn <> "" Then

            xMvtP0.Id = constYSOLDE0 & mId$(xIn, 1, 29)
            xMvtP0.Text = xIn
            dbMvtP0_Update xMvtP0
            
    End If
Loop

Close
Call lstErr_ChangeLastItem(Me.lstErr, Me.cmdContext, "Nb soldes : " & seq)

Exit Sub

Error_Handler:

blnError = True
Shell_MsgBox "me.cmdYSOLDE0_Import#  & error ", vbCritical, Me.Caption, False
Close

End Sub


Public Sub cmdYCOMPTE0_Import()
Dim wFileName As String
Dim xIn As String, X As String
Dim seq As Long, K As Integer

On Error GoTo Error_Handler

wFileName = paramSAB_Solde_Folder & Trim(txtSelect)
wFileName = fileName_Change(wFileName, constYSOLDE0, constYCOMPTE0)

Call lstErr_AddItem(lstErr, cmdContext, constYCOMPTE0 & ": "): DoEvents
recMvtP0_Init xMvtP0
xMvtP0.Method = constAddNew

seq = 0
Open wFileName For Input As #1

Do Until EOF(1)
    seq = seq + 1
    
    If seq Mod 500 = 0 Then Call lstErr_ChangeLastItem(Me.lstErr, Me.cmdContext, constYCOMPTE0 & " : " & seq)
    DoEvents
    Line Input #1, xIn
    If xIn <> "" Then

            xMvtP0.Id = constYCOMPTE0 & mId$(xIn, 1, 29)
            xMvtP0.Text = xIn
            dbMvtP0_Update xMvtP0
            
    End If
        
Loop

Close
Call lstErr_ChangeLastItem(Me.lstErr, Me.cmdContext, "Nb comptes : " & seq)

Exit Sub

Error_Handler:

blnError = True
Shell_MsgBox "me.cmdYCOMPTE0_Import#  & error ", vbCritical, Me.Caption, False
Close

End Sub


Private Sub mnuSelect_Print_Recap_Click()
'blnTotal = False
cmdPrint_Journal

End Sub

Private Sub txtSelect_GotFocus()
Call txt_GotFocus(txtSelect)

End Sub


Private Sub txtSelect_LostFocus()
Call txt_LostFocus(txtSelect)

End Sub



Public Sub fgSelect_ForeColor(lColor As Long)
For I = 0 To fgSelect_arrIndex
  fgSelect.Col = I: fgSelect.CellForeColor = lColor
Next I

End Sub


Public Sub fgSelect_Sort_Options()

fgSelect_Sort1_Old = -1
fgSelect_SortAD = 0
'If optOptions_SortUnit Then
'    If fgSelect_Sort1 <> 0 Then fgSelect_Sort1 = 0: fgSelect_Sort2 = 4: fgSelect_Sort
'End If

End Sub

Public Sub cmdPrint_Journal()
Me.Enabled = False

Msg = Space$(50)
Select Case fgSelect_Sort1
'    Case 0: prtSAB_Solde.prtSAB_Solde_Unit fgSelect, fgSelect_arrIndex, Me, blnTotal
'    Case 5: prtSAB_Solde.prtSAB_Solde_Devise fgSelect, fgSelect_arrIndex, Me, blnTotal
'    Case 6: prtSAB_Solde.prtSAB_Solde_Compte fgSelect, fgSelect_arrIndex, Me, blnTotal
End Select
Me.Enabled = True

End Sub

