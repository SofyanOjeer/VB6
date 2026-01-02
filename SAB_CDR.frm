VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmSAB_CDR 
   AutoRedraw      =   -1  'True
   Caption         =   "SAB_CDR :affichage des  paramètres"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13875
   Icon            =   "SAB_CDR.frx":0000
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
      TabCaption(0)   =   "Paramétrage SAB "
      TabPicture(0)   =   "SAB_CDR.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraSelect"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "?"
      TabPicture(1)   =   "SAB_CDR.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      Begin VB.Frame fraSelect 
         Height          =   8445
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   13560
         Begin VB.ListBox lstSelect 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   6060
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   10
            Top             =   2160
            Width           =   5295
         End
         Begin VB.ComboBox cboSelect_SQL 
            Height          =   315
            Left            =   6120
            Sorted          =   -1  'True
            TabIndex        =   9
            Text            =   "cboSelect_SQL"
            Top             =   240
            Width           =   3615
         End
         Begin VB.CommandButton cmdSelect_Ok 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Rechercher"
            Height          =   645
            Left            =   12360
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   120
            Width           =   1095
         End
         Begin VB.Frame fraSelect_Options 
            Height          =   525
            Left            =   11400
            TabIndex        =   6
            Top             =   240
            Width           =   795
         End
         Begin MSFlexGridLib.MSFlexGrid fgSelect 
            Height          =   6105
            Left            =   6120
            TabIndex        =   8
            Top             =   2280
            Width           =   7200
            _ExtentX        =   12700
            _ExtentY        =   10769
            _Version        =   393216
            Rows            =   1
            Cols            =   5
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
            FormatString    =   ">N°   |<Identifiant             |<Paramètres                        ||"
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
         Begin VB.Label lblSelect_Param 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   1335
            Left            =   9480
            TabIndex        =   13
            Top             =   840
            Width           =   3615
         End
         Begin VB.Label lblCRITAB01 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   1815
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   5295
         End
         Begin VB.Label lblSelect_Id 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   1335
            Left            =   6120
            TabIndex        =   11
            Top             =   840
            Width           =   3255
         End
      End
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   500
      Left            =   13320
      Picture         =   "SAB_CDR.frx":0044
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
Attribute VB_Name = "frmSAB_CDR"
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
Dim SAB_CDRAut As typeAuthorization
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

Dim fgSelect_D_FormatString As String

'______________________________________________________________________

Dim xZCRITAB0 As typeZCRITAB0
Dim arrZCRITAB0() As typeZCRITAB0, arrZCRITAB0_Nb As Long, arrZCRITAB0_Max As Long, arrZCRITAB0_Index As Long

Dim mCRITABNUM As Integer

'______________________________________________________________________

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
Private Sub fgSelect_Display()
Dim I As Long, X As String
Dim blnOk As Boolean, blnDisplay As Boolean
Dim wIndex As Long

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
fgSelect_Reset
cmdPrint.Enabled = False
lstSelect.Visible = False
lstSelect.Clear
lblSelect_Id = ""
lblSelect_Param = ""
fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
currentAction = "fgselect_Display"
mCRITABNUM = Val(mId$(cboSelect_SQL, 1, 1))

Select Case mCRITABNUM
    Case 2:   rsZBASTAB0_cboK2 5, lstSelect, ""
    Case 4:   rsZBASTAB0_cboK2 14, lstSelect, ""
               lblSelect_Id = "Etablissement    : ***" & vbCrLf _
                            & "Code produit     : **"
               lblSelect_Param = "Rubrique BDF 1: **" & vbCrLf _
                               & "Rubrique BDF 2: **" & vbCrLf _
                               & "Rubrique BDF 3: **"
                
    Case 5:   rsZBASTAB0_cboK2 58, lstSelect, ""
               lblSelect_Id = "Code opération : ***" & vbCrLf _
                              & "Nature opé     : ***" & vbCrLf _
                              & "Durée opération: ***"
               lblSelect_Param = "Rubrique BDF : **" & vbCrLf _
                               & "Autorisation : *"
    Case 6:   rsZPLAN0_cboK2 lstSelect
               lblSelect_Id = "N° plan        : ***" & vbCrLf _
                            & "Rubrique       : ******"
               lblSelect_Param = "Rubrique BDF 1: **" & vbCrLf _
                               & "Rubrique BDF 2: **" & vbCrLf _
                               & "Rubrique BDF 3: **"
                
End Select

For I = 1 To arrZCRITAB0_Nb
         
    xZCRITAB0 = arrZCRITAB0(I)
    If xZCRITAB0.CRITABNUM = mCRITABNUM Or mCRITABNUM = 0 Then
        fgSelect.Rows = fgSelect.Rows + 1
        fgSelect.Row = fgSelect.Rows - 1
        fgSelect_DisplayLine I
        Select Case mCRITABNUM
            Case 2: X = Trim(xZCRITAB0.CRITABARG)
                    Call lst_Scan(X, lstSelect)
                    If lstSelect.ListIndex >= 0 Then lstSelect.Selected(lstSelect.ListIndex) = True
            Case 4: X = Trim(mId$(xZCRITAB0.CRITABARG, 4, 3))
                    Call lst_Scan(X, lstSelect)
                    If lstSelect.ListIndex >= 0 Then lstSelect.Selected(lstSelect.ListIndex) = True
             Case 5: X = Trim(mId$(xZCRITAB0.CRITABARG, 1, 6))
                    Call lst_Scan(X, lstSelect)
                    If lstSelect.ListIndex >= 0 Then lstSelect.Selected(lstSelect.ListIndex) = True
              Case 6: X = Trim(mId$(xZCRITAB0.CRITABARG, 4, 6))
                    Call lst_Scan(X, lstSelect)
                    If lstSelect.ListIndex >= 0 Then lstSelect.Selected(lstSelect.ListIndex) = True
      End Select
    End If
Next I

Call lstErr_Clear(lstErr, cmdContext, "Nb enregistrements : " & fgSelect.Rows - 1): DoEvents
If fgSelect.Rows > 1 Then
    fgSelect_Sort1 = 0: fgSelect_Sort2 = 2: fgSelect_Sort
    If mCRITABNUM > 0 Then cmdPrint.Enabled = True
End If
lstSelect.Visible = True
fgSelect.Visible = True
lstSelect.SetFocus
If lstSelect.ListCount > 0 Then lstSelect.ListIndex = 0
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub

Public Sub fgSelect_DisplayLine(lIndex As Long)
On Error Resume Next
fgSelect.Col = 0: fgSelect.Text = xZCRITAB0.CRITABNUM
fgSelect.Col = 1: fgSelect.Text = xZCRITAB0.CRITABARG
fgSelect.Col = 2: fgSelect.Text = xZCRITAB0.CRITABDON

fgSelect.Col = fgSelect_arrIndex: fgSelect.Text = lIndex

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

Public Sub fgSelect_SortX(lK As Integer)
Dim I As Integer, X As String
For I = 1 To fgSelect.Rows - 1
    fgSelect.Row = I
    If lK = 2 Then
        fgSelect.Col = 2
        X = fgSelect.Text
    Else
        X = ""
    End If
    
    fgSelect.Col = 3
    X = X & Format$(Val(fgSelect.Text), "000000000000000.00")
    fgSelect.Col = fgSelect_arrIndex - 1
    fgSelect.Text = X
Next I


fgSelect_Sort1 = fgSelect_arrIndex - 1: fgSelect_Sort2 = fgSelect_arrIndex - 1
fgSelect_Sort
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
fgSelect_Display

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
lblSelect_Id.ForeColor = warnUsrColor '&H800000
lblSelect_Param.ForeColor = warnUsrColor
lblCRITAB01.ForeColor = warnUsrColor

cmdContext.Caption = constcmdRechercher: blnMsgBox_Quit = False
arrTag_Set False
lstErr.Visible = False
blncmdOk_Visible = False: blncmdSave_Visible = False
currentAction = ""

blnAuto = False
blnAuto_Ok = False

libRéférenceInterne = ""
cboSelect_SQL.ListIndex = 0
cmdSelect_SQL
lblCRITAB01_Display
blnControl = True
End Sub
Public Sub Form_Init()
Call lstErr_Clear(lstErr, cmdContext, "Initialisation ")
DoEvents

SSTab1.Tab = 0

blnControl = False

fgSelect_FormatString = fgSelect.FormatString
cmdSelect_Ok.Visible = False
fraSelect_Options.Visible = False
cboSelect_SQL.Clear
cboSelect_SQL.AddItem "0 - toutes les tables"
cboSelect_SQL.AddItem "2 - codes état"
cboSelect_SQL.AddItem "3 - rubriques BDF"
cboSelect_SQL.AddItem "4 - correspondance Produit/rubrique BDF"
cboSelect_SQL.AddItem "5 - correspondance Opération/rubrique BDF"
cboSelect_SQL.AddItem "6 - correspondance Rubrique Comptable/rubrique BDF"
cboSelect_SQL.AddItem "9 - collectif/nature relation"

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
cmdPrint_Ok
Me.Show

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdSelect_Ok_Click()
Dim blnOk As Boolean, Nb As Long


blnOk = fraSelect_Options.Enabled
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> SAB_CDR_cmdSelect_Ok ........"): DoEvents

fgSelect.Clear
If blnOk Then
    cmdSelect_Ok.Caption = "Options"
    cmdSelect_Ok.BackColor = &HFFFFFA   '&HC0FFFF
    fraSelect_Options.BackColor = &H8000000F
    Call usrColor_Container(fraSelect_Options, fraSelect_Options.BackColor)
    fraSelect_Options.Enabled = False
    cmdSelect_SQL
Else
    cmdSelect_Ok.Caption = constcmdRechercher
    cmdSelect_Ok.BackColor = &HC0FFC0
    fraSelect_Options.BackColor = &HFFFFFA    '&HC0FFFF
    Call usrColor_Container(fraSelect_Options, fraSelect_Options.BackColor)
    fraSelect_Options.Enabled = True
End If
Call lstErr_AddItem(lstErr, cmdContext, "< SAB_CDR_cmdSelect_Ok"): DoEvents
Me.Enabled = True: Me.MousePointer = 0


End Sub


Private Sub cmdSelect_SQL()
Dim V
Dim X As String
Dim xWhere As String, xAnd As String
Dim xDate10 As String
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL"
xWhere = ""
arrZCRITAB0_SQL xWhere


fgSelect_Display

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub


Private Sub arrZCRITAB0_SQL(xWhere As String)
Dim V
Dim X As String, xSql As String
On Error GoTo Error_Handler
ReDim arrZCRITAB0(101)
arrZCRITAB0_Max = 100: arrZCRITAB0_Nb = 0

Set rsSab = Nothing

xSql = "select * from " & paramIBM_Library_SAB & ".ZCRITAB0 " & xWhere & " order by CRITABNUM,CRITABARG"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    V = rsZCRITAB0_GetBuffer(rsSab, xZCRITAB0)

     If Not IsNull(V) Then
         MsgBox V, vbCritical, "frmSwift_Messages.fgselect_Display"
        '' Exit Sub
     Else
         arrZCRITAB0_Nb = arrZCRITAB0_Nb + 1
         If arrZCRITAB0_Nb > arrZCRITAB0_Max Then
             arrZCRITAB0_Max = arrZCRITAB0_Max + 50
             ReDim Preserve arrZCRITAB0(arrZCRITAB0_Max)
         End If
         
         arrZCRITAB0(arrZCRITAB0_Nb) = xZCRITAB0
         If xZCRITAB0.CRITABNUM = 1 Then arrZCRITAB0(0) = xZCRITAB0
    End If
    rsSab.MoveNext

Loop

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim K As Long
On Error Resume Next
If Y <= fgSelect.RowHeightMin Then
    Select Case fgSelect.Col
        Case 0: fgSelect_Sort1 = 0: fgSelect_Sort2 = 1: fgSelect_Sort
        Case 1:  fgSelect_Sort1 = 1: fgSelect_Sort2 = 1: fgSelect_Sort
        Case 2: fgSelect_Sort1 = 2: fgSelect_Sort2 = 2: fgSelect_Sort
        Case 3: fgSelect_Sort1 = 3: fgSelect_Sort2 = 3: fgSelect_Sort
        Case 4: fgSelect_Sort1 = 4: fgSelect_Sort2 = 4: fgSelect_Sort
       Case fgSelect_arrIndex:  fgSelect_SortX fgSelect_arrIndex
    End Select
Else
    If fgSelect.Rows > 1 Then
        Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
        fgSelect.Col = fgSelect_arrIndex:  K = CLng(fgSelect.Text)
        xZCRITAB0 = arrZCRITAB0(K)
       ' fgSelect_D_Display
        
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

Call BiaPgmAut_Init(mId$(Msg, 1, 12), SAB_CDRAut)

blnSetfocus = True
Form_Init


blnAuto = False


End Sub


Public Sub cmdContext_Return()
If SSTab1.Tab = 0 Then
'    cmdlstSourceScan_Click
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
    
For I = 1 To arrZCRITAB0_Nb
    fgSelect.Row = I
    Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
    fgSelect.Col = fgSelect_arrIndex:  K = CLng(fgSelect.Text)
    xZCRITAB0 = arrZCRITAB0(K)
    'prtSAB_CDR_Monitor xZCRITAB0
Next I

Me.Show

Me.Enabled = True: Me.MousePointer = 0



End Sub



Public Sub lblCRITAB01_Display()
Dim X As String
X = arrZCRITAB0(0).CRITABDON
lblCRITAB01 = "N° de la remise          : " & Val(mId$(X, 1, 6)) & vbCrLf _
            & "Seuil de déclaration     :" & Val(mId$(X, 8, 14)) / 100000 & vbCrLf _
            & "Compensation des comptes : " & mId$(X, 26, 1) & vbCrLf _
            & "Devise de déclaration    : " & mId$(X, 27, 3) & vbCrLf _
            & "Expression de la devise  : " & mId$(X, 30, 2) & vbCrLf _
            & "Gestion des collectifs   : " & mId$(X, 32, 1) & vbCrLf _
            & "Déclar par groupe Agences: " & mId$(X, 33, 1) & vbCrLf _
            & "Code remettant           : " & mId$(X, 34, 5) & vbCrLf
End Sub

Public Sub cmdPrint_Ok()
Dim K As Long
lstSelect.Visible = False
prtSAB_CDR_Open "Paramétrage SAB CDR"
For K = 1 To arrZCRITAB0_Nb
         
    xZCRITAB0 = arrZCRITAB0(K)
    If xZCRITAB0.CRITABNUM = mCRITABNUM Then

        prtSAB_CDR_NewLine
        XPrt.CurrentX = prtMinX + 50: XPrt.Print xZCRITAB0.CRITABNUM;
        XPrt.CurrentX = prtMinX + 500: XPrt.Print xZCRITAB0.CRITABARG;
        
        XPrt.CurrentX = prtMinX + 2000: XPrt.Print xZCRITAB0.CRITABDON;
        XPrt.CurrentX = prtMinX + 6000
          Select Case mCRITABNUM
              Case 2: X = Trim(xZCRITAB0.CRITABARG)
                      Call lst_Scan(X, lstSelect)
                        XPrt.Print lstSelect.Text;
              Case 4: X = Trim(mId$(xZCRITAB0.CRITABARG, 4, 3))
                      Call lst_Scan(X, lstSelect)
                      XPrt.Print lstSelect.Text;
               Case 5: X = Trim(mId$(xZCRITAB0.CRITABARG, 1, 6))
                      Call lst_Scan(X, lstSelect)
                      XPrt.Print lstSelect.Text;
                Case 6: X = Trim(mId$(xZCRITAB0.CRITABARG, 4, 6))
                      Call lst_Scan(X, lstSelect)
                      XPrt.Print lstSelect.Text;
        End Select
    End If
Next K
prtSAB_CDR_Close

lstSelect.Visible = True


End Sub
