VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmSysInvent 
   AutoRedraw      =   -1  'True
   Caption         =   "SysInvent : inventaire matériel"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13875
   Icon            =   "SysInvent.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9495
   ScaleWidth      =   13875
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   450
      Left            =   8280
      TabIndex        =   4
      Top             =   45
      Width           =   5055
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8895
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   15690
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Sélection"
      TabPicture(0)   =   "SysInvent.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraSelect"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Statistiques"
      TabPicture(1)   =   "SysInvent.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      Begin VB.Frame fraSelect 
         Height          =   8445
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   13560
         Begin MSFlexGridLib.MSFlexGrid fgSelect 
            Height          =   7545
            Left            =   120
            TabIndex        =   6
            Top             =   720
            Width           =   3720
            _ExtentX        =   6562
            _ExtentY        =   13309
            _Version        =   393216
            Rows            =   1
            Cols            =   6
            FixedCols       =   0
            RowHeightMin    =   300
            BackColor       =   16777210
            ForeColor       =   8388608
            BackColorFixed  =   16776921
            ForeColorFixed  =   -2147483641
            BackColorSel    =   12648384
            BackColorBkg    =   16777210
            WordWrap        =   -1  'True
            AllowBigSelection=   0   'False
            TextStyleFixed  =   4
            FocusRect       =   2
            HighLight       =   0
            GridLines       =   3
            GridLinesFixed  =   1
            AllowUserResizing=   3
            FormatString    =   " Utilisateur  |<PC          |<IP        |||"
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
         Begin VB.ComboBox cboSelect_SQL 
            Height          =   315
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   7
            Text            =   "cboSelect_SQL"
            Top             =   240
            Width           =   3615
         End
         Begin MSFlexGridLib.MSFlexGrid fgList 
            Height          =   7545
            Left            =   4080
            TabIndex        =   8
            Top             =   240
            Width           =   9240
            _ExtentX        =   16298
            _ExtentY        =   13309
            _Version        =   393216
            Rows            =   1
            Cols            =   6
            FixedCols       =   0
            RowHeightMin    =   300
            BackColor       =   16777210
            ForeColor       =   8388608
            BackColorFixed  =   16776921
            ForeColorFixed  =   -2147483641
            BackColorSel    =   12648384
            BackColorBkg    =   16777210
            WordWrap        =   -1  'True
            AllowBigSelection=   0   'False
            TextStyleFixed  =   4
            FocusRect       =   2
            HighLight       =   0
            GridLines       =   3
            GridLinesFixed  =   1
            AllowUserResizing=   3
            FormatString    =   " Utilisateur  |<PC          |<IP        |||"
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
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   500
      Left            =   13320
      Picture         =   "SysInvent.frx":0044
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
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Top             =   0
      Width           =   5535
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuContext 
      Caption         =   "mnuContext"
      Visible         =   0   'False
      Begin VB.Menu mnuContextOptions 
         Caption         =   "Options"
      End
      Begin VB.Menu mnuContext_x1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuContextAbandonner 
         Caption         =   "Abandonner"
      End
      Begin VB.Menu mnuContextQuitter 
         Caption         =   "Quitter"
      End
   End
   Begin VB.Menu mnuselect 
      Caption         =   "mnuSelect"
      Visible         =   0   'False
      Begin VB.Menu mnuSelect_Quit 
         Caption         =   "Abandonner"
      End
   End
   Begin VB.Menu mnuPrint0 
      Caption         =   "mnuPrint0"
      Visible         =   0   'False
      Begin VB.Menu mnuPrint0_All 
         Caption         =   "Imprimer TOUS les courriers"
      End
   End
End
Attribute VB_Name = "frmSysInvent"
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
Dim SysInvent_Aut As typeAuthorization
Dim blnTransaction As Boolean
Dim blnAuto As Boolean, blnAuto_Ok As Boolean
Dim wAmjMin As String, wAmjMax As String, wHmsMin As Long, wHmsMax As Long
Dim wAmjMin7 As Long, wAmjMax7 As Long


Dim blncmdOk_Visible As Boolean, blnErr As Boolean, blncmdSave_Visible As Boolean
Dim blnSetfocus As Boolean

Dim fgSelect_FormatString As String, fgSelect_K As Integer
Dim fgSelect_RowDisplay As Integer, fgSelect_RowClick As Integer, fgSelect_ColClick As Integer
Dim fgSelect_ColorClick As Long, fgSelect_ColorDisplay As Long
Dim fgSelect_Sort1 As Integer, fgSelect_Sort2 As Integer
Dim fgSelect_SortAD As Integer, fgSelect_Sort1_Old As Integer
Dim fgSelect_arrIndex As Integer
Dim blnfgSelect_DisplayLine As Boolean

Dim cmdSelect_SQL_K As String
Dim cmdSelect_Ok_Caption As String
'______________________________________________________________________

Dim fgList_FormatString As String, fgList_K As Integer
Dim fgList_RowDisplay As Integer, fgList_RowClick As Integer, fgList_ColClick As Integer
Dim fgList_ColorClick As Long, fgList_ColorDisplay As Long
Dim fgList_Sort1 As Integer, fgList_Sort2 As Integer
Dim fgList_SortAD As Integer, fgList_Sort1_Old As Integer
Dim fgList_arrIndex As Integer
Dim blnfgList_DisplayLine As Boolean
'______________________________________________________________________

Dim cnX As New ADODB.Connection
Dim rsX As New ADODB.Recordset
Dim arrS_System(200)   As typeS_System, arrS_System_Nb As Integer, arrS_System_Index As Integer
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
        For I = fgSelect_arrIndex To 0 Step -1
          fgSelect.Col = I: fgSelect.CellBackColor = lColor
        Next I
        fgSelect.LeftCol = 0
    End If
End If

End Sub
Public Sub fglist_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
mRow = fgList.Row

If lRow > 0 And lRow < fgList.Rows Then
    fgList.Row = lRow
    For I = 0 To fgList_arrIndex
        fgList.Col = I: fgList.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgList.Row = mRow
    If fgList.Row > 0 Then
        lRow = fgList.Row
        lColor_Old = fgList.CellBackColor
        For I = fgList_arrIndex To 0 Step -1
          fgList.Col = I: fgList.CellBackColor = lColor
        Next I
        fgList.LeftCol = 0
    End If
End If

End Sub

Private Sub fgSelect_Display_1()
Dim I As Long, X As String
Dim blnOk As Boolean, blnDisplay As Boolean
Dim wIndex As Long

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
fgSelect_Reset
fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
cmdPrint.Enabled = False
currentAction = "fgselect_Display"

For I = 1 To arrS_System_Nb

    
        fgSelect.Rows = fgSelect.Rows + 1
        fgSelect.Row = fgSelect.Rows - 1
        
        fgSelect_DisplayLine_1 I
Next I

Call lstErr_Clear(lstErr, cmdContext, "Nb enregistrements : " & fgSelect.Rows - 1): DoEvents
If fgSelect.Rows > 1 Then
'    fgSelect_Sort1 = 0: fgSelect_Sort2 = 2: fgSelect_Sort
    cmdPrint.Enabled = True
End If
fgSelect.Visible = True
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub
Private Sub fgSelect_Display_2(lX1 As String, lX2 As String)
Dim I As Long, X As String
Dim blnOk As Boolean, blnDisplay As Boolean
Dim wIndex As Long

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
fgSelect_Reset
fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
cmdPrint.Enabled = False
currentAction = "fgselect_Display_2"
For I = 1 To arrS_System_Nb
    arrS_System(I).blnOk = False
Next I

Set rsX = Nothing

X = "select PC from software where S_Name = '" & lX1 & "' and S_FileVers = '" & lX2 & "' order by PC"
Set rsX = cnX.Execute(X)

I = 1
Do While Not rsX.EOF

    Do
       X = Trim(rsX("PC"))
        If X = arrS_System(I).S_PC Then
            arrS_System(I).blnOk = True: Exit Do
        Else
            I = I + 1
        End If
    Loop Until I > arrS_System_Nb
    
    rsX.MoveNext
Loop
fgSelect_Display_1

Call lstErr_Clear(lstErr, cmdContext, "Nb enregistrements : " & fgSelect.Rows - 1): DoEvents
If fgSelect.Rows > 1 Then
    fgSelect_Sort1 = 3: fgSelect_Sort2 = 3: fgSelect_SortX 3
    cmdPrint.Enabled = True
End If
fgSelect.Visible = True
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub

Public Sub fgSelect_DisplayLine_1(lIndex As Long)
Dim X As String
On Error Resume Next
fgSelect.Col = 0: fgSelect.Text = arrS_System(lIndex).S_UserName
fgSelect.CellForeColor = IIf(arrS_System(lIndex).blnOk, vbBlue, vbRed)
fgSelect.Col = 1: fgSelect.Text = arrS_System(lIndex).S_PC
fgSelect.CellForeColor = IIf(arrS_System(lIndex).blnOk, vbBlue, vbRed)
 
fgSelect.Col = fgSelect_arrIndex - 2: fgSelect.Text = IIf(arrS_System(lIndex).blnOk, 0, 1)
fgSelect.Col = fgSelect_arrIndex: fgSelect.Text = lIndex

End Sub


Private Sub lstSelect_Load_1()
Dim I As Long, xSql As String
Dim blnOk As Boolean, blnDisplay As Boolean
Dim xWhere As String

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
cmdPrint.Enabled = False
currentAction = "lstSelect_Load_1"
cmdSelect_Ok_Caption = "Lancer la requête"
fgSelect_FormatString = " Utilisateur  |<PC          |<IP        ||"
fgSelect_Display_1



Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub

Private Sub lstSelect_Load_2()
Dim I As Long, xSql As String
Dim blnOk As Boolean, blnDisplay As Boolean
Dim xWhere As String

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
cmdPrint.Enabled = False
currentAction = "lstSelect_Load_2"
fgList_Display_2
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub
Public Sub fgSelect_Reset()
fgSelect.Clear
fgSelect_Sort1 = 0: fgSelect_Sort2 = 0
fgSelect_Sort1_Old = -1
fgSelect_RowDisplay = 0: fgSelect_RowClick = 0
fgSelect_arrIndex = fgSelect.Cols - 1
blnfgSelect_DisplayLine = False
fgSelect_SortAD = 6
fgSelect.LeftCol = 0

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
Public Sub fglist_Sort()
If fgList.Rows > 1 Then
    fgList.Row = 1
    fgList.RowSel = fgList.Rows - 1
    
    If fgList_Sort1_Old = fgList_Sort1 Then
        If fgList_SortAD = 5 Then
            fgList_SortAD = 6
        Else
            fgList_SortAD = 5
        End If
    Else
        fgList_SortAD = 5
    End If
    fgList_Sort1_Old = fgList_Sort1
    
    fgList.Col = fgList_Sort1
    fgList.ColSel = fgList_Sort2
    fgList.Sort = fgList_SortAD
End If

End Sub

Public Sub fgSelect_SortX(lK As Integer)
Dim I As Integer, X As String
Dim wIndex As Integer
For I = 1 To fgSelect.Rows - 1
    fgSelect.Row = I
    fgSelect.Col = fgSelect_arrIndex
    wIndex = Val(fgSelect.Text)
    Select Case lK
        Case 3
                fgSelect.Col = 3: X = Format$(Val(fgSelect.Text), "0000")
                fgSelect.Col = 0: X = X & Trim(fgSelect.Text)
    End Select
    fgSelect.Col = fgSelect_arrIndex - 1
    fgSelect.Text = X
Next I


fgSelect_Sort1 = fgSelect_arrIndex - 1: fgSelect_Sort2 = fgSelect_arrIndex - 1
fgSelect_Sort
End Sub
Private Sub fgList_Display_1(lPC As String)
Dim I As Long, X As String
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgList.Visible = False
fglist_Reset
fgList.Rows = 1
fgList.FormatString = "<Nom               |<Version     |<Path                                            ||"
cmdPrint.Enabled = False
currentAction = "fgList_Display_1"

Set rsX = Nothing

X = "select S_Name,S_FileVers,S_FullName from software where PC = '" & lPC & "' order by S_FullName"
Set rsX = cnX.Execute(X)
Do While Not rsX.EOF
        fgList.Rows = fgList.Rows + 1
        fgList.Row = fgList.Rows - 1

    fgList.Col = 0: fgList.Text = Trim(rsX("S_Name"))
    fgList.Col = 1: fgList.Text = Trim(rsX("S_FileVers"))
    fgList.Col = 2: fgList.Text = Trim(rsX("S_FullName"))
    rsX.MoveNext
Loop


fgList.Visible = True
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction

fgList.Visible = True


End Sub

Private Sub fgList_Display_2()
Dim I As Long, X As String
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgList.Visible = False
fglist_Reset
fgList.Rows = 1
fgList.FormatString = "Nom               |Version     | Path                                            ||"
cmdPrint.Enabled = False
currentAction = "fgList_Display_1"

Set rsX = Nothing

X = "select distinct(S_Name),S_FileVers,S_FullName from software  order by S_FullName"
Set rsX = cnX.Execute(X)
Do While Not rsX.EOF
        fgList.Rows = fgList.Rows + 1
        fgList.Row = fgList.Rows - 1

    fgList.Col = 0: fgList.Text = Trim(rsX("S_Name"))
    fgList.Col = 1: fgList.Text = Trim(rsX("S_FileVers"))
    fgList.Col = 2: fgList.Text = Trim(rsX("S_FullName"))
    rsX.MoveNext
Loop


fgList.Visible = True
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction

fgList.Visible = True


End Sub


Public Sub fglist_Reset()
fgList.Clear
fgList_Sort1 = 0: fgList_Sort2 = 0
fgList_Sort1_Old = -1
fgList_RowDisplay = 0: fgList_RowClick = 0
fgList_arrIndex = fgList.Cols - 1
blnfgList_DisplayLine = False
fgList_SortAD = 6
fgList.LeftCol = 0

End Sub


'---------------------------------------------------------
Public Sub arrTag_Set(ByVal B As Boolean)
'---------------------------------------------------------
For I = 0 To arrTagNb: arrTag(I) = B: Next I
End Sub


Public Sub cmdContext_Quit()
Unload Me
End Sub




Private Sub cboSelect_SQL_Click()
cmdSelect_SQL_K = mId$(cboSelect_SQL, 1, 1)
If blnControl Then
    Me.Enabled = False: Me.MousePointer = vbHourglass
    Select Case cmdSelect_SQL_K
        Case "1": lstSelect_Load_1
        Case "2": lstSelect_Load_2
    End Select
    Me.Enabled = True: Me.MousePointer = 0
End If
End Sub


Private Sub cmdContext_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdContext

End Sub

Private Sub cmdPrint_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdPrint

End Sub

'---------------------------------------------------------
Public Sub cmdReset()
'---------------------------------------------------------
Dim I As Integer

blnControl = False
usrColor_Set

cmdContext.Caption = constcmdRechercher: blnMsgBox_Quit = False
arrTag_Set False
lstErr.Visible = False
blncmdOk_Visible = False: blncmdSave_Visible = False
currentAction = ""

blnAuto = False
blnAuto_Ok = False

libRéférenceInterne = ""
cboSelect_SQL.ListIndex = 0
fgSelect.Visible = False
cboSelect_SQL.ListIndex = 1
blnControl = True
cboSelect_SQL.ListIndex = 0
End Sub
Public Sub Form_Init()
Call lstErr_Clear(lstErr, cmdContext, "Initialisation ")
DoEvents

SSTab1.Tab = 0

blnControl = False

fgSelect.Visible = False
fgSelect_FormatString = fgSelect.FormatString
cboSelect_SQL.Clear
cboSelect_SQL.AddItem "1 - Utilisateur => Logiciels"
cboSelect_SQL.AddItem "2 - Logiciels => Utilisateurs"



cnX.CursorLocation = ADODB.CursorLocationEnum.adUseClient
cnX.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" _
& "SERVER=192.168.168.16;" _
& "UID=sysinvent_adm;" _
& "PWD=Manage;" _
& "OPTION=" & 1 + 2 + 8 + 32 + 2048 + 16384

cnX.Open
cnX.Execute ("USE sysinvent") 'select database
X = "select * from system where Sys_Info = 'UserName' order by PC"
Set rsX = cnX.Execute(X)
arrS_System_Nb = 0
Do While Not rsX.EOF
    arrS_System_Nb = arrS_System_Nb + 1
    arrS_System(arrS_System_Nb).S_PC = rsX("PC")
    arrS_System(arrS_System_Nb).S_UserName = rsX("Sys_Valeur")
    arrS_System(arrS_System_Nb).blnOk = True
    rsX.MoveNext
Loop

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

Me.Enabled = False: Me.MousePointer = vbHourglass
    Select Case cmdSelect_SQL_K
    '    Case "2": cmdPrint_Ok_2
    '    Case "3": cmdPrint_Ok_3
    End Select

Me.Show

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub fgList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim K As Long
Dim X1 As String, X2 As String
Me.Enabled = False
On Error Resume Next
If Y <= fgList.RowHeightMin Then
        Select Case fgList.Col
            Case 0: fgList_Sort1 = 0: fgList_Sort2 = 0: fglist_Sort
            Case 1:  fgList_Sort1 = 1: fgList_Sort2 = 1: fglist_Sort
            Case 2: fgList_Sort1 = 2: fgList_Sort2 = 2: fglist_Sort
        End Select
Else
    If cmdSelect_SQL_K = "2" And fgList.Rows > 1 Then
        Me.Enabled = False: Me.MousePointer = vbHourglass
        fgList.Col = 0
        Call fglist_Color(fgList_RowClick, MouseMoveUsr.BackColor, fgList_ColorClick)
        fgList.Col = 0: X1 = Trim(fgList.Text)
        fgList.Col = 1: X2 = Trim(fgList.Text)
        
        fgSelect_Display_2 X1, X2
        Me.Enabled = True: Me.MousePointer = 0
   End If
End If
Me.Enabled = True

End Sub


Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim K As Long
Me.Enabled = False
On Error Resume Next
If Y <= fgSelect.RowHeightMin Then
        Select Case fgSelect.Col
            Case 0: fgSelect_Sort1 = 0: fgSelect_Sort2 = 0: fgSelect_Sort
            Case 1:  fgSelect_Sort1 = 1: fgSelect_Sort2 = 1: fgSelect_Sort
            Case 2: fgSelect_Sort1 = 2: fgSelect_Sort2 = 2: fgSelect_Sort
           Case fgSelect_arrIndex:  fgSelect_SortX fgSelect_arrIndex
        End Select
Else
    If cmdSelect_SQL_K = "1" And fgSelect.Rows > 1 Then
        Me.Enabled = False: Me.MousePointer = vbHourglass
        fgSelect.Col = fgSelect_arrIndex:  arrS_System_Index = CLng(fgSelect.Text)
        Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
       
        fgList_Display_1 arrS_System(arrS_System_Index).S_PC
        Me.Enabled = True: Me.MousePointer = 0

   End If
End If
Me.Enabled = True
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

Private Sub Form_Unload(Cancel As Integer)
cnX.Close
End Sub

Private Sub mnuContextAbandonner_Click()
cmdContext_Quit
End Sub

Private Sub mnuContextQuitter_Click()
Unload Me
End Sub

Public Sub Msg_Rcv(Msg As String)
'---------------------------------------------------------
Dim meUnit As typeUnit, X As String
mWindowState = Me.WindowState
If Me.WindowState <> frmUsr_Windowstate Then Me.WindowState = frmUsr_Windowstate

Call BiaPgmAut_Init(mId$(Msg, 1, 12), SysInvent_Aut)

blnSetfocus = True
Form_Init
blnAuto = False

Select Case UCase$(Trim(mId$(Msg, 1, 12)))

    Case Else: blnAuto = False
                
                    
End Select


End Sub


Public Sub cmdContext_Return()
If SSTab1.Tab = 0 Then
'    If fraUpdate.Visible _
'   And fraUpdate_B.Enabled _
'    And cmdUpdate_Ok.Enabled Then cmdUpdate_Ok_Click: Exit Sub
Else
    If currentAction = "" Then
        If SSTab1.Tab > 0 Then
            SSTab1.Tab = 0
        Else
           'SendKeys "{TAB}"
           ' cmdSelect_Click
        End If
    End If
End If
End Sub









Private Sub mnuPrint0_All_Click()
Dim I As Long, K As Long
Me.Enabled = False: Me.MousePointer = vbHourglass
    

Me.Show

Me.Enabled = True: Me.MousePointer = 0



End Sub




