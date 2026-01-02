VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSAB_Ordonnanceur 
   AutoRedraw      =   -1  'True
   Caption         =   "SAB_Ordonnanceur"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9090
   Icon            =   "SAB_Ordonnanceur.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6210
   ScaleWidth      =   9090
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   450
      Left            =   4350
      TabIndex        =   4
      Top             =   -30
      Width           =   4260
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5445
      Left            =   -45
      TabIndex        =   2
      Top             =   495
      Width           =   9090
      _ExtentX        =   16034
      _ExtentY        =   9604
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Ordonnanceur"
      TabPicture(0)   =   "SAB_Ordonnanceur.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraTab0"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "...."
      TabPicture(1)   =   "SAB_Ordonnanceur.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraTAb1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Maintenance informatique"
      TabPicture(2)   =   "SAB_Ordonnanceur.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraInfo"
      Tab(2).ControlCount=   1
      Begin VB.Frame fraInfo 
         Height          =   4905
         Left            =   -74910
         TabIndex        =   7
         Top             =   480
         Width           =   8820
         Begin VB.CommandButton cmdYBAST049 
            BackColor       =   &H00C0C0FF&
            Caption         =   "Chargement de la table ZBASTAB0 : 049 =>Info.mdb"
            Height          =   810
            Left            =   210
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   225
            Width           =   2400
         End
      End
      Begin VB.Frame fraTAb1 
         Height          =   4905
         Left            =   -74940
         TabIndex        =   6
         Top             =   390
         Width           =   8940
      End
      Begin VB.Frame fraTab0 
         Height          =   4875
         Left            =   135
         TabIndex        =   3
         Top             =   360
         Width           =   8850
         Begin VB.CommandButton cmdSelect 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Chargement de l'ordonnanceur SAB"
            Height          =   435
            Left            =   75
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   165
            Width           =   3015
         End
         Begin VB.Frame fraContextOptions 
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
            Height          =   3975
            Left            =   4830
            TabIndex        =   9
            Top             =   675
            Width           =   3795
         End
         Begin MSFlexGridLib.MSFlexGrid fgSelect 
            Height          =   4155
            Left            =   75
            TabIndex        =   5
            Top             =   585
            Width           =   8595
            _ExtentX        =   15161
            _ExtentY        =   7329
            _Version        =   393216
            Rows            =   1
            Cols            =   8
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
            FormatString    =   $"SAB_Ordonnanceur.frx":035E
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
      Left            =   8600
      Picture         =   "SAB_Ordonnanceur.frx":040C
      Style           =   1  'Graphical
      TabIndex        =   0
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
         Caption         =   "Imprimer l'ordonnanceur"
      End
   End
End
Attribute VB_Name = "frmSAB_Ordonnanceur"
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
Dim SAB_ORDONNAN_Aut As typeAuthorization
Dim blnAuto As Boolean, blnError As Boolean

Dim fgSelect_FormatString As String, fgSelect_K As Integer
Dim fgSelect_RowDisplay As Integer, fgSelect_RowClick As Integer, fgSelect_ColClick As Integer
Dim fgSelect_ColorClick As Long, fgSelect_ColorDisplay As Long
Dim fgSelect_Sort1 As Integer, fgSelect_Sort2 As Integer
Dim fgSelect_SortAD As Integer, fgSelect_Sort1_Old As Integer
Dim fgSelect_arrIndex As Integer
Dim blnfgSelect_DisplayLine As Boolean

Dim meYBASTAB0 As typeYBASTAB0, xYBASTAB0 As typeYBASTAB0
Dim meYCOMTAC0 As typeYCOMTAC0, xYCOMTAC0 As typeYCOMTAC0
Dim meYCOMEXP0 As typeYCOMEXP0, xYCOMEXP0 As typeYCOMEXP0

Dim xMVTP0 As typeMvtP0, memMvtp0 As typeMvtP0
Dim seq As Long
Dim xElpKMInfo As typeElpKMInfo, meElpKMIndex As typeElpKMIndex

Dim mCOMTACNUM As String

Private Sub cmdYBAST049_Import()

Dim blnUpdate As Boolean
Dim kIn As Integer, seq As Long
On Error GoTo Error_Handle
Dim xIn As String, X As String
Dim xIn2 As String

On Error GoTo Error_Handle
X = MsgBox("Voulez-vous vraiment mettre à jour YBAST049 (delete/AddNew)  ?", vbQuestion + vbYesNo, Me.Caption)
If X = vbNo Then Exit Sub

MDB.Execute "delete * from elpkmInfo where ElpKMSrc_Id = 2000"
MDB.Execute "delete * from ElpKMIndex where ElpKMSrc_Id = 2000"

Call lstErr_AddItem(lstErr, cmdContext, "cmdYBAST049_Import : " & Time): DoEvents
xIn = paramTemp_Folder & "FTP\YBASTAB0_049.txt"
Open Trim(xIn) For Input As #1

seq = 0
mdbElpKMInfo.tableElpKMInfo_Open

recElpKMInfo_Init xElpKMInfo
xElpKMInfo.Method = constAddNew
xElpKMInfo.ElpKMSrc_Id = 2000
xElpKMInfo.Pass = 1000

mdbElpKMIndex.tableElpKMIndex_Open

recElpKMIndex_Init meElpKMIndex
meElpKMIndex.Method = constAddNew
meElpKMIndex.Classe = 1000
Call lstErr_AddItem(lstErr, cmdContext, "......"): DoEvents

'''Exit Sub
Do Until EOF(1)
    seq = seq + 1
    If seq Mod 100 = 0 Then Call lstErr_ChangeLastItem(Me.lstErr, Me.cmdContext, "cmdSAB_ZBAST049_Write : " & seq)
    DoEvents
    Line Input #1, xIn
    If xIn <> "" Then
            kIn = 0
            MsgTxt = Space$(34) & xIn
            MsgTxtIndex = 0
            srvYBASTAB0_GetBuffer meYBASTAB0

            xElpKMInfo.Id = Trim(meYBASTAB0.BASTABARG)
            xIn2 = Trim(mId$(meYBASTAB0.BASTABDON, 1, 40))
            xElpKMInfo.Description = xIn2
            
            xElpKMInfo.Memo = xIn
            dbElpKMInfo_Update xElpKMInfo
            
            meElpKMIndex.ElpKMSrc_Id = xElpKMInfo.ElpKMSrc_Id
            xIn2 = Text_LCase(xIn2)
            blnUpdate = True
            
            Do
                X = Text_KeyWord(xIn2, kIn, False)
            
                If X <> "" Then
                    meElpKMIndex.Id = X

                    meElpKMIndex.Method = "Seek="
                    If tableElpKMIndex_Read(meElpKMIndex) = 0 Then
                        meElpKMIndex.Method = constUpdate
                        meElpKMIndex.Memo = meElpKMIndex.Memo & xElpKMInfo.Id
                    Else
                        meElpKMIndex.Method = constAddNew
                        meElpKMIndex.Memo = xElpKMInfo.Id
                    End If
                    
                    dbElpKMIndex_Update meElpKMIndex

                Else
                    blnUpdate = False
                End If
                
            Loop While blnUpdate
        End If
        
Loop

Call lstErr_Clear(Me.lstErr, Me.cmdContext, "cmdYBAST049_import fin : " & seq)

Close

Exit Sub

Error_Handle:
 MsgBox "erreur : cmdYBAST049_Import  " & xIn, vbCritical, Error
Close


End Sub


Public Sub cmdYCOMTAC0_Import()
Dim mFileName As String
Dim xIn As String, X As String


On Error GoTo Error_Handler

mFileName = Trim(paramTemp_Folder & "FTP\YCOMTAC0.txt")
Call lstErr_AddItem(lstErr, cmdContext, mFileName & ": début"): DoEvents
recMvtP0_Init xMVTP0
xMVTP0.Method = constAddNew

seq = 0
Open mFileName For Input As #1

Do Until EOF(1)
    seq = seq + 1
    
    If seq Mod 10 = 0 Then Call lstErr_ChangeLastItem(Me.lstErr, Me.cmdContext, mFileName & " : " & seq)
    DoEvents
    Line Input #1, xIn
    If xIn <> "" Then
            MsgTxt = Space$(34) & xIn
            MsgTxtIndex = 0
            srvYCOMTAC0_GetBuffer meYCOMTAC0

            xMVTP0.Id = Format$(meYCOMTAC0.COMTACETA, "0000") & meYCOMTAC0.COMTACTRA & meYCOMTAC0.COMTACOPT _
                      & Format$(seq, "000000") & "YCOMTAC0"
            xMVTP0.Text = xIn
            dbMvtP0_Update xMVTP0
            
    End If
        
Loop

Close
Call lstErr_ChangeLastItem(Me.lstErr, Me.cmdContext, "cmdYCOMTAC0_Import : fin " & seq)

Exit Sub

Error_Handler:

blnError = True
Shell_MsgBox "frmSAB_Ordonnanceur.cmdYCOMTAC0_Import#  & error ", vbCritical, Me.Caption, True
Close

End Sub



Public Sub cmdYCOMEXP0_Import()
Dim mFileName As String
Dim xIn As String, X As String


On Error GoTo Error_Handler

mFileName = Trim(paramTemp_Folder & "FTP\YCOMEXP0.txt")
Call lstErr_AddItem(lstErr, cmdContext, mFileName & ": début"): DoEvents
recMvtP0_Init xMVTP0
xMVTP0.Method = constAddNew

Open mFileName For Input As #1

Do Until EOF(1)
    seq = seq + 1
    
    If seq Mod 10 = 0 Then Call lstErr_ChangeLastItem(Me.lstErr, Me.cmdContext, mFileName & " : " & seq)
    DoEvents
    Line Input #1, xIn
    If xIn <> "" Then
            MsgTxt = Space$(34) & xIn
            MsgTxtIndex = 0
            srvYCOMEXP0_GetBuffer meYCOMEXP0

            xMVTP0.Id = Format$(meYCOMEXP0.COMEXPETA, "0000") & meYCOMEXP0.COMEXPTRA & meYCOMEXP0.COMEXPOPT _
                      & Format$(seq, "000000") & "YCOMEXP0"
            xMVTP0.Text = xIn
            dbMvtP0_Update xMVTP0
            
    End If
        
Loop

Close
Call lstErr_ChangeLastItem(Me.lstErr, Me.cmdContext, "cmdYCOMEXP0_Import : fin " & seq)

Exit Sub

Error_Handler:

blnError = True
Shell_MsgBox "frmSAB_Ordonnanceur.cmdYCOMEXP0_Import#  & error ", vbCritical, Me.Caption, True
Close

End Sub


Private Sub fgSelect_Display()
SSTab1.Tab = 0

fgSelect_Reset

fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
fgSelect.Visible = False

xMVTP0.Id = ""
xMVTP0.Method = "MoveFirst"
intReturn = tableMvtP0_Read(xMVTP0)
xMVTP0.Method = "Seek>"

Do
    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1

    fgSelect_DisplayLine
     
    intReturn = tableMvtP0_Read(xMVTP0)
    
Loop Until intReturn <> 0


fgSelect.Visible = True

fgSelect_Sort
End Sub

Public Sub fgSelect_DisplayLine()
On Error Resume Next

    MsgTxt = Space$(34) & xMVTP0.Text
    MsgTxtIndex = 0
    If mId$(xMVTP0.Id, 20, 8) = "YCOMTAC0" Then
        srvYCOMTAC0_GetBuffer xYCOMTAC0
        fgSelect.Col = 0: fgSelect.Text = xYCOMTAC0.COMTACPER
        fgSelect.Col = 1: fgSelect.Text = xYCOMTAC0.COMTACTRA
        mCOMTACNUM = Format$(xYCOMTAC0.COMTACNUM, "00000")
        fgSelect.Col = 2: fgSelect.Text = mCOMTACNUM
        fgSelect.Col = 3: fgSelect.Text = xYCOMTAC0.COMTACOPT
        xElpKMInfo.Id = xYCOMTAC0.COMTACOPT
        fgSelect.Col = 4
        If tableElpKMInfo_Read(xElpKMInfo) = 0 Then
             fgSelect.Text = xElpKMInfo.Description
        Else
            fgSelect.Text = "????????????????"
        End If
        
        fgSelect.Col = 6: fgSelect.Text = "YCOMTAC0"
     Else
        srvYCOMEXP0_GetBuffer xYCOMEXP0
        If xYCOMEXP0.COMEXPOPT = xYCOMTAC0.COMTACOPT Then
            fgSelect.Col = 2: fgSelect.Text = mCOMTACNUM & "_"   '''''comtac précédent
        Else
            fgSelect.Col = 2: fgSelect.Text = "?????"
        End If
        fgSelect.Col = 0: fgSelect.Text = ""
        fgSelect.Col = 1: fgSelect.Text = xYCOMEXP0.COMEXPTRA
        fgSelect.Col = 3: fgSelect.Text = xYCOMEXP0.COMEXPOPT
        fgSelect.Col = 4: fgSelect.Text = xYCOMEXP0.COMEXPARG & " ...." & xYCOMEXP0.COMEXPDON
       fgSelect.Col = 6: fgSelect.Text = "YCOMEXP0"
     
        fgSelect_ForeColor warnUsrColor
     End If
End Sub

Public Sub fgSelect_Sort()
On Error Resume Next
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
    fgSelect.Col = lK
    X = Format$(Val(fgSelect.Text), "0000000")
    fgSelect.Col = fgSelect_arrIndex - 1
    fgSelect.Text = X
    'Select Case lK
    '    Case 1, 2: fgSelect.Text = X
    'End Select
Next I


fgSelect_Sort1 = fgSelect_arrIndex - 1: fgSelect_Sort2 = fgSelect_arrIndex - 1
fgSelect_Sort
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

Call BiaPgmAut_Init("SAB_Ordonnanceur", SAB_ORDONNAN_Aut)

'blnSetfocus = True
Form_Init


End Sub


Public Sub Form_Init()
Me.Enabled = False
Call lstErr_Clear(lstErr, cmdPrint, "Initialisation ")
DoEvents
lstErr.Visible = True
If Not IsNull(param_Init) Then
    MsgBox "paramétrage inconsistent", vbCritical, "frmSAB_Ordonnanceur.param_init"
    fraTab0.Enabled = False
End If

    blnControl = False
    fgSelect_FormatString = fgSelect.FormatString
    fgSelect.Enabled = True
    fraTAb1.Visible = SAB_ORDONNAN_Aut.Xspécial
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
fraContextOptions.Visible = False
blnControl = True

End Sub



Public Function param_Init()
Dim K As Integer, K1 As Integer, X As String

Dim V
param_Init = Null
On Error GoTo Table_Error


Exit Function

Table_Error:
param_Init = V
Shell_MsgBox "frmSAB_Ordonnanceur.param_init#  " & recElpTable.Id & " : " & recElpTable.K1 & " : " & recElpTable.K2 & V, vbCritical, Me.Caption, True
Exit Function

Memo_Error:
param_Init = "Memo"
Shell_MsgBox "frmSAB_Ordonnanceur.param_init#  " & recElpTable.Id & " : " & recElpTable.K1 & " : " & recElpTable.K2 & " manque memo", vbCritical, Me.Caption, True
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

Private Sub cmdPrint_Click()

Msg = Space$(50)
Select Case SSTab1.Tab
    Case 0: Me.PopupMenu mnuPrint0, vbPopupMenuLeftButton
End Select

End Sub

Private Sub cmdSelect_Click()
Dim mFileName As String
Dim seq As Long, lenX As Long
Dim K As Integer, I As Integer
Dim X As String
Dim xIn As String
Dim blnOk As Boolean, blnDevises As Boolean, blnEONIA As Boolean
Dim mAAMMJJ As String * 8, mHHMMSS As String * 6
Dim xAAMMJJ As String * 8, xHHMMSS As String * 6
Dim mNature As String, xCodeDevise As String, xCodeTaux As String, xVal As Double
On Error Resume Next

Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "Ordonnanceur_Import"): DoEvents

mdbMvtP0.tableMvtP0_Close
MDB.Execute "delete * from mvtp0"
mdbMvtP0.tableMvtP0_Open

mdbElpKMInfo.tableElpKMInfo_Open

recElpKMInfo_Init xElpKMInfo
xElpKMInfo.Method = "Seek="
xElpKMInfo.ElpKMSrc_Id = 2000

cmdYCOMTAC0_Import
cmdYCOMEXP0_Import


fgSelect_Display


Me.Enabled = True: Me.MousePointer = 0


End Sub

Private Sub cmdYBAST049_Click()
cmdYBAST049_Import
End Sub

Private Sub fgSelect_Click()
fgSelect.LeftCol = 0

End Sub

Private Sub fgSelect_LeaveCell()
On Error Resume Next
'fgSelect.CellBackColor = &HE0E0E0
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

If fraContextOptions.Visible Then fraContextOptions.Visible = False: Exit Sub

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
On Error Resume Next
If Y <= fgSelect.RowHeightMin Then
    Select Case fgSelect.Col
        Case 0: fgSelect_Sort1 = 0: fgSelect_Sort2 = 2: fgSelect_Sort
        Case 1:  fgSelect_Sort1 = 1: fgSelect_Sort2 = 2: fgSelect_Sort
        Case 2: fgSelect_Sort1 = 2: fgSelect_Sort2 = 2: fgSelect_Sort
        Case 3: fgSelect_Sort1 = 3: fgSelect_Sort2 = 3: fgSelect_Sort
        Case 4: fgSelect_Sort1 = 4: fgSelect_Sort2 = 4: fgSelect_Sort
        Case fgSelect_arrIndex:  fgSelect_SortX fgSelect_arrIndex
    End Select
Else
    If fgSelect.Rows > 1 Then
        Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
        fgSelect.Col = fgSelect_arrIndex
   End If
End If
End Sub

Public Sub fgSelect_Reset()
fgSelect.Clear
fgSelect_Sort1 = 1: fgSelect_Sort2 = 2
fgSelect_Sort1_Old = -1
fgSelect_RowDisplay = 0: fgSelect_RowClick = 0
fgSelect_arrIndex = 7
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



Public Sub fgSelect_ForeColor(lColor As Long)
For I = 0 To fgSelect_arrIndex
  fgSelect.Col = I: fgSelect.CellForeColor = lColor
Next I

End Sub

Private Sub mnuSelect_Print_Click()
Me.Enabled = False

Msg = Space$(50)
prtSAB_Ordonnanceur.prtSAB_Ordonnanceur fgSelect
Me.Enabled = True

End Sub


