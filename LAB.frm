VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmLAB 
   AutoRedraw      =   -1  'True
   Caption         =   "LAB : lutte anti blanchiment"
   ClientHeight    =   9870
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13875
   Icon            =   "LAB.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9870
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
      Height          =   9255
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   16325
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "World-Check"
      TabPicture(0)   =   "LAB.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraSelect"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "...."
      TabPicture(1)   =   "LAB.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      Begin VB.ListBox lstW 
         Height          =   255
         Left            =   7200
         Sorted          =   -1  'True
         TabIndex        =   13
         Top             =   -300
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Frame fraSelect 
         Height          =   8445
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   13560
         Begin VB.Frame fraSelect_Update 
            BackColor       =   &H00808080&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   6495
            Left            =   4320
            TabIndex        =   19
            Top             =   1560
            Visible         =   0   'False
            Width           =   9015
            Begin VB.CommandButton cmdSelect_Update_Quit 
               BackColor       =   &H00C0C0FF&
               Caption         =   "Abandonner"
               Height          =   525
               Left            =   4320
               Style           =   1  'Graphical
               TabIndex        =   32
               Top             =   5640
               Width           =   1575
            End
            Begin VB.Frame fraSelect_Update_A 
               BackColor       =   &H00E0FFFF&
               Height          =   1455
               Left            =   120
               TabIndex        =   23
               Top             =   120
               Width           =   8775
               Begin VB.TextBox txtUpdate_WC_Id 
                  Height          =   285
                  Left            =   2040
                  TabIndex        =   27
                  Top             =   240
                  Width           =   1455
               End
               Begin VB.TextBox txtUpdate_WC_LastName 
                  Height          =   285
                  Left            =   2040
                  TabIndex        =   26
                  Top             =   600
                  Width           =   6015
               End
               Begin VB.TextBox txtUpdate_WC_FirstName 
                  Height          =   285
                  Left            =   2040
                  TabIndex        =   25
                  Top             =   960
                  Width           =   6015
               End
               Begin VB.TextBox txtUpdate_WC_Sta 
                  Height          =   285
                  Left            =   7560
                  TabIndex        =   24
                  Top             =   240
                  Width           =   495
               End
               Begin VB.Label lblUpdate_WC_Id 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "Identifiant"
                  Height          =   255
                  Left            =   360
                  TabIndex        =   31
                  Top             =   240
                  Width           =   975
               End
               Begin VB.Label lblUpdate_WC_LastName 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "Nom"
                  Height          =   255
                  Left            =   360
                  TabIndex        =   30
                  Top             =   720
                  Width           =   735
               End
               Begin VB.Label lblUpdate_WC_FirstName 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "Prénom"
                  Height          =   255
                  Left            =   360
                  TabIndex        =   29
                  Top             =   1080
                  Width           =   735
               End
               Begin VB.Label lblUpdate_WC_Sta 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "Etat"
                  Height          =   255
                  Left            =   6360
                  TabIndex        =   28
                  Top             =   240
                  Width           =   975
               End
            End
            Begin VB.Frame fraSelect_Update_B 
               BackColor       =   &H00D0D0D0&
               Height          =   4695
               Left            =   120
               TabIndex        =   20
               Top             =   1680
               Width           =   8775
               Begin VB.TextBox txtUpdate_WC_Memo 
                  Height          =   3375
                  Left            =   2040
                  Locked          =   -1  'True
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   33
                  Top             =   360
                  Width           =   6015
               End
               Begin VB.CommandButton cmdSelect_Update_Ok 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "Enregistrer"
                  Height          =   525
                  Left            =   6480
                  Style           =   1  'Graphical
                  TabIndex        =   22
                  Top             =   4080
                  Width           =   1575
               End
               Begin VB.CommandButton cmdSelect_Update_Annuler 
                  BackColor       =   &H000000FF&
                  Caption         =   "Annuler définitivement"
                  Height          =   525
                  Left            =   2040
                  Style           =   1  'Graphical
                  TabIndex        =   21
                  Top             =   4080
                  Width           =   1575
               End
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Frame1"
            Height          =   135
            Left            =   8760
            TabIndex        =   18
            Top             =   2880
            Width           =   15
         End
         Begin VB.ListBox List1 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5520
            Left            =   240
            Sorted          =   -1  'True
            TabIndex        =   16
            Top             =   1680
            Visible         =   0   'False
            Width           =   1935
         End
         Begin MSFlexGridLib.MSFlexGrid fgSelect 
            Height          =   6825
            Left            =   120
            TabIndex        =   8
            Top             =   1560
            Visible         =   0   'False
            Width           =   13440
            _ExtentX        =   23707
            _ExtentY        =   12039
            _Version        =   393216
            Rows            =   1
            Cols            =   12
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
            FormatString    =   $"LAB.frx":0044
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
            Left            =   9240
            Sorted          =   -1  'True
            TabIndex        =   9
            Text            =   "cboSelect_SQL"
            Top             =   240
            Width           =   3615
         End
         Begin VB.CommandButton cmdSelect_Ok 
            BackColor       =   &H00E0FFFF&
            Caption         =   "Rechercher"
            Height          =   645
            Left            =   10200
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   840
            Width           =   1815
         End
         Begin VB.Frame fraSelect_Options_1 
            Height          =   1125
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   8115
            Begin VB.ComboBox txtSelect_WC_Sta 
               Height          =   315
               Left            =   2280
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   15
               Top             =   720
               Width           =   2775
            End
            Begin VB.CheckBox chkSelect_WC_UpdD 
               Caption         =   "Période de création"
               Height          =   255
               Left            =   120
               TabIndex        =   12
               Top             =   120
               Width           =   1815
            End
            Begin VB.TextBox txtSelect_WC_LastName 
               Height          =   285
               Left            =   5760
               TabIndex        =   11
               Top             =   720
               Width           =   1935
            End
            Begin MSComCtl2.DTPicker txtSelect_WC_UpdD 
               Height          =   300
               Left            =   480
               TabIndex        =   10
               Top             =   360
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
               Format          =   20643843
               CurrentDate     =   38699.44875
               MaxDate         =   401768
               MinDate         =   36526.4425347222
            End
            Begin MSComCtl2.DTPicker txtSelect_WC_UpdD_Max 
               Height          =   300
               Left            =   480
               TabIndex        =   14
               Top             =   720
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
               Format          =   20643843
               CurrentDate     =   38699.44875
               MaxDate         =   401768
               MinDate         =   36526.4425347222
            End
            Begin VB.Label lblSelect_WC_LastName 
               Alignment       =   2  'Center
               Caption         =   "ID ou NOM  (partiel)"
               Height          =   255
               Left            =   5760
               TabIndex        =   17
               Top             =   240
               Width           =   1935
            End
         End
      End
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   500
      Left            =   13320
      Picture         =   "LAB.frx":00ED
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
   Begin VB.Menu mnuPrint0 
      Caption         =   "mnuPrint0"
      Visible         =   0   'False
      Begin VB.Menu mnuPrint0_Liste 
         Caption         =   "Imprimer la liste"
      End
   End
End
Attribute VB_Name = "frmLAB"
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
Dim LAB_Aut As typeAuthorization
Dim blnTransaction As Boolean
Dim blnAuto As Boolean, blnAuto_Ok As Boolean
Dim wAmjMin As String, wAmjMax As String, wHmsMin As Long, wHmsMax As Long
Dim wAmjMin7 As Long, wAmjMax7 As Long
Dim rsWCX As New ADODB.Recordset


Dim blncmdOk_Visible As Boolean, blnErr As Boolean, blncmdSave_Visible As Boolean
Dim blnSetfocus As Boolean

Dim fgSelect_FormatString As String, fgSelect_K As Integer
Dim fgSelect_RowDisplay As Integer, fgSelect_RowClick As Integer, fgSelect_ColClick As Integer
Dim fgSelect_ColorClick As Long, fgSelect_ColorDisplay As Long
Dim fgSelect_Sort1 As Integer, fgSelect_Sort2 As Integer
Dim fgSelect_SortAD As Integer, fgSelect_Sort1_Old As Integer
Dim fgSelect_arrIndex As Integer
Dim blnfgSelect_DisplayLine As Boolean
Dim cmdSelect_Ok_Caption As String
Dim cmdSelect_SQL_K As String
Dim xWC_Data As typeWC_Data, meWC_Data As typeWC_Data
Dim newWC_Data As typeWC_Data, oldWC_Data As typeWC_Data
Dim arrWC_Data() As typeWC_Data, arrWC_Data_Nb As Long, arrWC_Data_Max As Long, arrWC_Data_Index As Long
Dim selWC_Data() As typeWC_Data, selWC_Data_Nb As Long, selWC_Data_Max As Long, selWC_Data_Index As Long

Dim objFolder As Folder, objFiles As Files
Dim fsoFile As File
Dim paramWorldCheck_Path_Input As String, paramWorldCheck_Path_Download As String, paramWorldCheck_Path_log As String
Dim arrWC_List_I() As String, arrWC_List_Nb As Long
Private Sub arrWC_Data_SQL(xWhere As String)
Dim V
Dim X As String, xSql As String
On Error GoTo Error_Handler
ReDim arrWC_Data(501)
arrWC_Data_Max = 500: arrWC_Data_Nb = 0

Set rsWC = Nothing

xSql = "select * from  WC_Data " & xWhere
Set rsWC = cnWC.Execute(xSql)

Do While Not rsWC.EOF
    V = rsWC_Data_GetBuffer(rsWC, xWC_Data)

     If Not IsNull(V) Then
         MsgBox V, vbCritical, "frmSwift_Messages.fgselect_Display"
        '' Exit Sub
     Else
         arrWC_Data_Nb = arrWC_Data_Nb + 1
         If arrWC_Data_Nb > arrWC_Data_Max Then
             arrWC_Data_Max = arrWC_Data_Max + 50
             ReDim Preserve arrWC_Data(arrWC_Data_Max)
         End If
         
         arrWC_Data(arrWC_Data_Nb) = xWC_Data
    End If
    rsWC.MoveNext

Loop

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

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
Private Sub fgSelect_Display()
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

For I = 1 To arrWC_Data_Nb

        xWC_Data = arrWC_Data(I)
    
        fgSelect.Rows = fgSelect.Rows + 1
        fgSelect.Row = fgSelect.Rows - 1
        fgSelect_DisplayLine I
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
cmdSelect_Ok.Caption = cmdSelect_Ok_Caption
cmdSelect_Ok.Visible = True
fraSelect_Options_1.Visible = True
fraSelect_Options_1.Enabled = True
chkSelect_WC_UpdD.Enabled = True
chkSelect_WC_UpdD = "0"
txtSelect_WC_Sta.Enabled = True
txtSelect_WC_LastName.Enabled = True
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction
End Sub

Public Sub fgSelect_DisplayLine(lIndex As Long)
Dim X As String, wColor As Long

On Error Resume Next
     Select Case xWC_Data.WC_Sta
        Case Is = "0": wColor = vbBlue
        Case Is = "1": wColor = vbRed
        Case Is = "2": wColor = vbMagenta
        Case Else: wColor = vbBlack
    End Select
fgSelect.Col = 0: fgSelect.Text = xWC_Data.WC_Id
fgSelect.CellForeColor = wColor
fgSelect.Col = 1: fgSelect.Text = dateImp10(xWC_Data.WC_UpdD) & " " & xWC_Data.WC_UpdH
fgSelect.CellForeColor = wColor
fgSelect.Col = 2: fgSelect.Text = xWC_Data.WC_Sta
fgSelect.CellForeColor = wColor
fgSelect.Col = 3: fgSelect.Text = xWC_Data.WC_LastName
fgSelect.CellForeColor = wColor
fgSelect.Col = 4: fgSelect.Text = xWC_Data.WC_FirstName
fgSelect.CellForeColor = wColor
fgSelect.Col = 5: fgSelect.Text = xWC_Data.WC_Memo
fgSelect.CellForeColor = wColor

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
Dim wIndex As Integer
For I = 1 To fgSelect.Rows - 1
    fgSelect.Row = I
    fgSelect.Col = fgSelect_arrIndex
    wIndex = Val(fgSelect.Text)
    Select Case lK
        Case 0: X = Format$(arrWC_Data(wIndex).WC_Id, "000000000")
        Case 1: X = Format$(arrWC_Data(wIndex).WC_UpdD, "000000000")
    End Select
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
Select Case SSTab1.Tab
    Case Is = 1
    Case Else
    
        If fraSelect_Update.Visible Then fraSelect_Update.Visible = False: Exit Sub
        If fgSelect.Visible Then fgSelect.Visible = False: cmdSelect_Ok.Caption = "Extraire les factures": Exit Sub
        Unload Me
End Select
End Sub





Private Sub cboSelect_SQL_Click()
cmdSelect_SQL_K = mId$(cboSelect_SQL, 1, 1)
If blnControl Then
    Me.Enabled = False: Me.MousePointer = vbHourglass
    fraSelect_Options_1.Visible = False
    fraSelect_Update.Visible = False
    Select Case cmdSelect_SQL_K
        Case "1": lstSelect_Load_1
    End Select
    Me.Enabled = True: Me.MousePointer = 0
End If
End Sub


Private Sub chkSelect_WC_UpdD_Click()
Select Case chkSelect_WC_UpdD
    Case Is = "1"
        txtSelect_WC_UpdD.Visible = True
        txtSelect_WC_UpdD_Max.Visible = True
    Case Is = "7"
        txtSelect_WC_UpdD.Visible = True
        txtSelect_WC_UpdD_Max.Visible = False
   Case Else
    txtSelect_WC_UpdD.Visible = False
    txtSelect_WC_UpdD_Max.Visible = False
End Select


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
fraSelect_Options_1.Visible = False
fraSelect_Update.Visible = False
cmdSelect_Ok.Caption = "Extraire les mouvements"

libRéférenceInterne = ""
If cboSelect_SQL.ListCount > 0 Then cboSelect_SQL.ListIndex = 0
lstSelect_Load_1
Call DTPicker_Set(txtSelect_WC_UpdD, YBIATAB0_DATE_CPT_JP0)
Call DTPicker_Set(txtSelect_WC_UpdD_Max, YBIATAB0_DATE_CPT_J)
fgSelect.Visible = False  'True
'cmdSelect_Ok_Click




blnControl = True
End Sub
Public Sub Form_Init()
Dim V, X As String
Dim xZBASTAB0 As typeZBASTAB0
Dim K As Integer

Call lstErr_Clear(lstErr, cmdContext, "Initialisation ")
DoEvents


blnControl = False

fgSelect_FormatString = fgSelect.FormatString
fraSelect.Enabled = LAB_Aut.Consulter
cmdSelect_Ok.Visible = False
fraSelect_Options_1.Visible = False
txtSelect_WC_UpdD.Visible = False
txtSelect_WC_UpdD_Max.Visible = False
cboSelect_SQL.Clear
If LAB_Aut.Consulter Then
    cboSelect_SQL.AddItem "1 - Liste WC"
End If
If LAB_Aut.Xspécial Then
    cboSelect_SQL.AddItem "6 - Filtre & import"
End If


'_____________________________________________________________________________
txtSelect_WC_Sta.Clear
txtSelect_WC_Sta.AddItem " "
txtSelect_WC_Sta.AddItem "0 - Others "
txtSelect_WC_Sta.AddItem "1 - Sanctions(création) "
txtSelect_WC_Sta.AddItem "2 - Sanctions(modification)"
txtSelect_WC_Sta.ListIndex = 0

paramWorldCheck_Path_Input = paramServer("\\SWIFT\" & paramEnvironnement & "\World-Check\Input\")
paramWorldCheck_Path_Download = paramServer("\\SWIFT\" & paramEnvironnement & "\World-Check\Download\")
paramWorldCheck_Path_log = paramServer("\\SWIFT\" & paramEnvironnement & "\World-Check\Log\")
MsgBox "c:\temp\"
paramWorldCheck_Path_Input = "c:\temp\World-Check\Input\"
paramWorldCheck_Path_Download = "c:\temp\World-Check\Download\"
paramWorldCheck_Path_log = "c:\temp\World-Check\Log\"
'_____________________________________________________________________________
' chargement des listes WC à ignorer
'_____________________________________________________________________________
X = "select count(*) as Tally from ElpTable  " _
    & " where id = 'LAB' and K1 = 'WC_List_I' "
Set rsMDB = cnMDB.Execute(X)
arrWC_List_Nb = rsMDB("Tally")

ReDim arrWC_List_I(arrWC_List_Nb + 1)
X = "select * from  ElpTable " _
    & " where id = 'LAB' and K1 = 'WC_List_I' "
    
Set rsMDB = cnMDB.Execute(X)
K = 0
Do While Not rsMDB.EOF
        K = K + 1
        V = rsMDB("Memo")
        If Not IsNull(V) And V <> "" Then
            arrWC_List_I(K) = Trim(V)
        Else
            arrWC_List_I(K) = Trim(rsMDB("K2"))
        End If
    rsMDB.MoveNext
Loop

'_____________________________________________________________________________
cmdReset




SSTab1.Visible = True



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
Select Case SSTab1.Tab
End Select
Me.Show

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdSelect_Ok_Click()
Dim blnOk As Boolean, Nb As Long

Me.Enabled = False: Me.MousePointer = vbHourglass
blnOk = Not fgSelect.Visible
Call lstErr_Clear(lstErr, cmdContext, "> SAB_CDR_cmdSelect_Ok ........"): DoEvents
cmdSelect_Ok.Visible = False
fraSelect_Update.Visible = False
fraSelect_Options_1.Enabled = False
fgSelect.Clear
DoEvents
If blnOk Then
    cmdSelect_Ok.Caption = "Modifier les options"
    cmdSelect_Ok.BackColor = &HFFFFFA   '&HC0FFFF
    fraSelect_Options_1.BackColor = &H8000000F
    Call usrColor_Container(fraSelect_Options_1, fraSelect_Options_1.BackColor)
    Select Case cmdSelect_SQL_K
        Case "1": cmdSelect_SQL_1
        Case "6": cmdSelect_SQL_6

    End Select

    fgSelect.Enabled = True
Else
    cmdSelect_Ok.Caption = cmdSelect_Ok_Caption
    cmdSelect_Ok.BackColor = &HE0FFFF
    fraSelect_Options_1.BackColor = &HE0FFFF  ' &HFFFFFA    '&HC0FFFF
    Call usrColor_Container(fraSelect_Options_1, fraSelect_Options_1.BackColor)
    fgSelect.Visible = False
    fgSelect.Enabled = False
    fraSelect_Options_1.Enabled = True

End If
Call lstErr_AddItem(lstErr, cmdContext, "< SAB_CDR_cmdSelect_Ok"): DoEvents
Me.Enabled = True: Me.MousePointer = 0
cmdSelect_Ok.Visible = True

End Sub


Private Sub cmdSelect_SQL_1()
Dim V, X As String
Dim xSql As String, K As Long
Dim xWhere As String, xAnd As String
On Error GoTo Error_Handler

Set rsWC = Nothing
Call lstErr_Clear(lstErr, cmdContext, "cmdSelect_SQL"): DoEvents

currentAction = "cmdSelect_SQL"
xWhere = ""
X = Trim(txtSelect_WC_LastName)
If X <> "" Then
    If IsNumeric(X) Then
        xWhere = " and   WC_ID = " & Val(X)
    Else
        xWhere = " and   WC_LastName like '%" & X & "%'"
    End If
End If

X = Trim(mId$(txtSelect_WC_Sta, 1, 1))
If X <> "" Then xWhere = xWhere & " and   WC_Sta = '" & X & "'"

Set rsWC = Nothing
Call DTPicker_Control(txtSelect_WC_UpdD, wAmjMin)
Call DTPicker_Control(txtSelect_WC_UpdD_Max, wAmjMax)

If chkSelect_WC_UpdD = "1" Then
    xWhere = xWhere & " and   WC_UpdD >= '" & wAmjMin & "'" _
                    & " and   WC_UpdD <= '" & wAmjMax & "'"
End If
If xWhere = "" Then
    Call lstErr_AddItem(lstErr, cmdContext, "? préciser les critères de recherche"): DoEvents
    fgSelect.Visible = True
Else
    Mid$(xWhere, 1, 6) = " where"
    arrWC_Data_SQL xWhere & " order by WC_LastName,WC_LastName"
        
    fgSelect_Display
End If
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Public Sub cmdSelect_SQL_6_BlackListed(lFileName As String)
Dim V
Dim X1 As String, xIn As String
Dim Nb1 As Long, Nb2 As Long, Nb3 As Long, Nb4 As Long
Dim K As Integer, K1 As Integer, K2 As Integer, K3 As Integer, K4 As Integer, xSql As String
Dim wFile_No As Integer
Dim rsWC_Data As typeWC_Data
Dim wFirstName As String, wLastName As String
Dim blnBlackListed As Boolean, blnUpd As Boolean
Dim xFileName As String
Dim xFileName_New As String
Dim X As String, wBlackListed As String
Dim wDSYS_Time As String
Dim wMotif As String, wSUBJECT As String

On Error GoTo Error_Handler
wDSYS_Time = DSYS_Time

Call lstErr_Clear(lstErr, cmdContext, lFileName): DoEvents

xFileName = paramWorldCheck_Path_Download & lFileName
Open xFileName For Input As #1
Open paramWorldCheck_Path_Input & "World-check-BIA_BlackListed.csv" For Output As #2
Open paramWorldCheck_Path_Input & "World-check-BIA_OTHERS.csv" For Output As #3
xFileName_New = paramWorldCheck_Path_log & "BIA_BlackListed_New_" & wDSYS_Time & ".txt"
Open xFileName_New For Output As #4

xIn = ""
Nb1 = 0: Nb4 = 0
rsWC_Data_Init rsWC_Data
rsWC_Data.WC_UpdD = DSys
rsWC_Data.WC_UpdH = time_Hms
Call lstErr_AddItem(lstErr, cmdContext, "Open"): DoEvents

Do Until EOF(1)
    Do
        X1 = Input(1, #1)
        xIn = xIn & X1
    Loop Until X1 = vbLf ' Asc10
    
    If Nb1 = 0 Then
        Nb1 = 1
        Print #2, xIn;
        Nb2 = 1
        Print #3, xIn;
        Nb2 = 1
    Else
    
        Nb1 = Nb1 + 1
        K = InStr(xIn, vbTab)
        rsWC_Data.WC_Id = Val(mId$(xIn, 1, K - 1))
        K1 = InStr(K + 1, xIn, vbTab)
        wLastName = Trim(mId$(xIn, K + 1, K1 - K - 1))
        K2 = InStr(K1 + 1, xIn, vbTab)
        wFirstName = Trim(mId$(xIn, K1 + 1, K2 - K1 - 1))
        K = 3
        Do
           K1 = InStr(K2 + 1, xIn, vbTab)
           K = K + 1
           If K = 22 Then
                wBlackListed = Trim(mId$(xIn, K2 + 1, K1 - K2 - 1))
                If wBlackListed <> "" Then
                    K3 = InStr(1, wBlackListed, "~")
                    If K3 = 0 Then
                        For K4 = 1 To arrWC_List_Nb
                            If wBlackListed = arrWC_List_I(K4) Then wBlackListed = "": Exit For
                        Next K4
                    
                    End If
                End If
                Exit Do
            End If
           K2 = K1
        Loop Until K1 = 0
  
       If wBlackListed <> "" Then
       'If InStr(xIn, "[SANCTIONS]") > 0 Then
            wFile_No = 2
            Nb2 = Nb2 + 1
            rsWC_Data.WC_Sta = "1"
            blnBlackListed = True
        Else
            wFile_No = 3
            Nb3 = Nb3 + 1
            rsWC_Data.WC_Sta = "0"
            blnBlackListed = False
        End If
    
        
        xSql = "select WC_sta,WC_Memo from WC_Data where WC_Id = " & rsWC_Data.WC_Id
        
        Set rsWC = cnWC.Execute(xSql)
        If rsWC.EOF Then
            rsWC_Data.WC_LastName = wLastName
            rsWC_Data.WC_FirstName = wFirstName
            rsWC_Data.WC_Memo = xIn
            V = adoWC_Data_AddNew(rsWC, rsWC_Data)
            Print #wFile_No, xIn;
            If blnBlackListed Then Nb4 = Nb4 + 1: Print #4, rsWC_Data.WC_Id; ";"; wLastName; ";"; wBlackListed
            If Not IsNull(V) Then GoTo Error_MsgBox

        Else
            blnUpd = False
            rsWC_Data.WC_Sta = rsWC("WC_Sta")
            rsWC_Data.WC_Memo = rsWC("WC_Memo")
            If xIn <> rsWC_Data.WC_Memo Then
                blnUpd = True
                If rsWC_Data.WC_Sta = "1" Then rsWC_Data.WC_Sta = "2"
            End If
           If blnBlackListed And rsWC_Data.WC_Sta = "0" Then blnUpd = True: rsWC_Data.WC_Sta = "1"
'$$$$$$$$ gérer levée des sanctions
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
            If blnUpd Then
                rsWC_Data.WC_LastName = wLastName
                rsWC_Data.WC_FirstName = wFirstName
                rsWC_Data.WC_Memo = xIn
                V = adoWC_Data_Update(rsWC, rsWC_Data)
                Print #wFile_No, xIn;
                If Not IsNull(V) Then GoTo Error_MsgBox

            End If
       End If
    End If
    If Nb1 Mod 1000 = 0 Then Call lstErr_ChangeLastItem(lstErr, cmdContext, Nb1 & " > " & Nb2 & " + " & Nb3): DoEvents

    xIn = ""
Loop

Call lstErr_AddItem(lstErr, cmdContext, Nb1 & " = " & Nb2 & " + " & Nb3): DoEvents
If Nb1 <> Nb2 + Nb3 Then MsgBox "ERREUR nb1 <> nb2 +nb3"
Close
Kill xFileName
If Nb4 > 0 Then cmdSendMail xFileName_New, "WorldCheck", "LAB", Nb4 & "mises à jour World-Check-BlackListed du " & wDSYS_Time

Exit Sub

Error_Handler:

    V = Error

Error_MsgBox:

    wMotif = "BIA_System > frmLAB > cmdSelect_SQL_6 : " & Nb1 & " = " & Nb2 & " + " & Nb3
    wSUBJECT = rsWC_Data.WC_Id & " : " & V
    If blnAuto Then
        Call SMS_MONITOR("ALERTE", wSUBJECT, wMotif)
    Else
        MsgBox wSUBJECT, vbCritical, wMotif
    End If
Close

End Sub

Private Sub cmdSelect_Update_Quit_Click()
fraSelect_Update.Visible = False

End Sub

Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim K As Long
Me.Enabled = False
On Error Resume Next
If Y <= fgSelect.RowHeightMin Then
        Select Case fgSelect.Col
            Case 0: fgSelect_Sort1 = 0: fgSelect_Sort2 = 0: fgSelect_SortX 0
            Case 1:  fgSelect_Sort1 = 1: fgSelect_Sort2 = 1: fgSelect_SortX 1
            Case 2: fgSelect_Sort1 = 2: fgSelect_Sort2 = 2: fgSelect_Sort
            Case 3: fgSelect_Sort1 = 3: fgSelect_Sort2 = 3: fgSelect_Sort
            Case 4: fgSelect_Sort1 = 4: fgSelect_Sort2 = 4: fgSelect_Sort
           Case fgSelect_arrIndex:  fgSelect_SortX fgSelect_arrIndex
        End Select
Else
    If fgSelect.Rows > 1 Then
        fgSelect.Col = fgSelect_arrIndex:  arrWC_Data_Index = CLng(fgSelect.Text)
        Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
        xWC_Data = arrWC_Data(arrWC_Data_Index)
        oldWC_Data = xWC_Data
        fraSelect_Display
   End If
End If
Me.Enabled = True: Me.MousePointer = 0
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
cnWC.Close
Set cnWC = Nothing

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

Call BiaPgmAut_Init(mId$(Msg, 1, 12), LAB_Aut)

blnSetfocus = True
Form_Init
WC_Data_Open
blnAuto = False

Select Case UCase$(Trim(mId$(Msg, 1, 12)))
    Case "@AUTO_WC": blnAuto = True
                        If Not IsEmpty(XPrt) Then Set Xprt_Previous = XPrt
                        Printer_PDF
'                         Call cbo_Scan("3 -", cboDétail_SQL)
 '                        cmdDétail_Ok_Click

                         Unload Me

    Case Else: blnAuto = False
End Select


End Sub


Public Sub cmdContext_Return()
Select Case SSTab1.Tab
    Case Is = 0
    Case Else
        If currentAction = "" Then
            If SSTab1.Tab > 0 Then
                SSTab1.Tab = 0
            Else
               'SendKeys "{TAB}"
               ' cmdSelect_Click
            End If
        End If
End Select
End Sub









Public Sub cmdSendMail(lFileName As String, lFrom As String, lRecipient As String, lSubject As String)
Dim wSendMail As typeSendMail
Dim bgColor As String
wSendMail.FromDisplayName = lFrom
wSendMail.RecipientDisplayName = lRecipient

bgColor = "cyan"
wSendMail.Subject = lSubject
wSendMail.Attachment = lFileName
wSendMail.Message = "<body bgcolor=" & Asc34 & bgColor & Asc34 & ">" _
                    & "<FONT face=" & Asc34 & prtFontName_Arial & Asc34 & ">" _
                    & htmlFontColor("BLUE") & "<B><CENTER>" & "voir pièce jointe" _
                    & "<BR>"

wSendMail.AsHTML = True

srvSendMail.Monitor wSendMail

End Sub

Public Sub fraSelect_Display()
Dim V
Dim X As String, X1 As String
fraSelect_Update.Visible = True
fraSelect_Update_A.Enabled = False
fraSelect_Update_B.Enabled = True
cmdSelect_Update_Annuler.Visible = LAB_Aut.Xspécial
cmdSelect_Update_Ok.Visible = LAB_Aut.Xspécial
txtUpdate_WC_Id = xWC_Data.WC_Id
txtUpdate_WC_Sta = xWC_Data.WC_Sta
txtUpdate_WC_LastName = xWC_Data.WC_LastName: txtUpdate_WC_LastName.BackColor = txtUsr.BackColor
txtUpdate_WC_FirstName = xWC_Data.WC_FirstName: txtUpdate_WC_FirstName.BackColor = txtUsr.BackColor
txtUpdate_WC_Memo = xWC_Data.WC_Memo: txtUpdate_WC_Memo.BackColor = txtUsr.BackColor

Call lstErr_Clear(lstErr, cmdContext, ">Affichage du détail"): DoEvents
End Sub





Private Sub mnuPrint0_Liste_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
'cmdPrint_Facture

Me.Show

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub txtSelect_WC_LastName_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub



Public Sub cmdSelect_SQL_6()
Dim Nb As Integer
Dim wMotif As String, wSUBJECT As String

'---------------------------------
'traitement précédent terminé ?
'---------------------------------
Set objFolder = msFileSystem.GetFolder(paramWorldCheck_Path_Input)
Set objFiles = objFolder.Files
Nb = 0
For Each fsoFile In objFiles
    Nb = Nb + 1
Next

If Nb > 0 Then
    wMotif = "BIA_System > frmLAB > cmdSelect_SQL_6"
    wSUBJECT = Nb & " fichiers non importés dans SAFEWATCH : " & paramWorldCheck_Path_Input
    If blnAuto Then
        Call SMS_MONITOR("ALERTE", wSUBJECT, wMotif)
    Else
        MsgBox wSUBJECT, vbCritical, wMotif
    End If
    Exit Sub
End If

'---------------------------------
' fichiers à traiter ? !!! 1 seul à la fois
'---------------------------------

Set objFolder = msFileSystem.GetFolder(paramWorldCheck_Path_Download)
Set objFiles = objFolder.Files
Nb = 0
For Each fsoFile In objFiles
    If Nb = 0 Then
        cmdSelect_SQL_6_BlackListed fsoFile.Name
    Else
        Nb = Nb + 1
    End If
Next
If Nb > 0 Then
    wMotif = "BIA_System > frmLAB > cmdSelect_SQL_6"
    wSUBJECT = Nb & " fichiers téléchargés non traités dans : " & paramWorldCheck_Path_Download
    If blnAuto Then
        Call SMS_MONITOR("ALERTE", wSUBJECT, wMotif)
    Else
        MsgBox wSUBJECT, vbCritical, wMotif
    End If
    Exit Sub
End If

End Sub
 
