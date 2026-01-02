VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmBia_NET_CMD 
   AutoRedraw      =   -1  'True
   Caption         =   "Bia_NET_CMD"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13875
   Icon            =   "BIA_NET_CMD.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9495
   ScaleWidth      =   13875
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   255
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
      TabPicture(0)   =   "BIA_NET_CMD.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraSelect"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "."
      TabPicture(1)   =   "BIA_NET_CMD.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fgSelect"
      Tab(1).ControlCount=   1
      Begin VB.Frame fraSelect 
         Height          =   8445
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   13560
         Begin VB.CommandButton cmdUpdate_Quit 
            BackColor       =   &H00C0C0FF&
            Caption         =   "Abandonner"
            Height          =   972
            Left            =   8400
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   1320
            Width           =   1575
         End
         Begin VB.CommandButton cmdUpdate_Ok 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Enregistrer"
            Height          =   1005
            Left            =   8400
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   240
            Width           =   1575
         End
         Begin VB.ComboBox cboSelect_SQL 
            Height          =   315
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   6
            Text            =   "cboSelect_SQL"
            Top             =   240
            Width           =   3615
         End
         Begin VB.Label lblMonitor_MONAMJ 
            Caption         =   "date"
            Height          =   252
            Left            =   2400
            TabIndex        =   15
            Top             =   1680
            Width           =   3372
         End
         Begin VB.Label lblMonitor_MONSTATUS 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "statut"
            Height          =   252
            Left            =   1440
            TabIndex        =   14
            Top             =   1680
            Width           =   732
         End
         Begin VB.Label lblService_MONAMJ 
            Caption         =   "date"
            Height          =   252
            Left            =   2400
            TabIndex        =   13
            Top             =   960
            Width           =   2892
         End
         Begin VB.Label lblService_MONSTATUS 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "statut"
            Height          =   252
            Left            =   1440
            TabIndex        =   12
            Top             =   960
            Width           =   732
         End
         Begin VB.Label lblMonitor 
            Caption         =   "commande"
            Height          =   252
            Left            =   120
            TabIndex        =   10
            Top             =   1680
            Width           =   852
         End
         Begin VB.Label lblService 
            Caption         =   "service"
            Height          =   252
            Left            =   120
            TabIndex        =   9
            Top             =   960
            Width           =   1092
         End
      End
      Begin MSFlexGridLib.MSFlexGrid fgSelect 
         Height          =   1908
         Left            =   -74040
         TabIndex        =   11
         Top             =   1200
         Visible         =   0   'False
         Width           =   9120
         _ExtentX        =   16087
         _ExtentY        =   3360
         _Version        =   393216
         Rows            =   1
         Cols            =   13
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
         FormatString    =   "       ||                           |||"
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
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   500
      Left            =   13320
      Picture         =   "BIA_NET_CMD.frx":0044
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
         Caption         =   "Imprimer"
      End
   End
End
Attribute VB_Name = "frmBia_NET_CMD"
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
Dim x As String, I As Integer, Msg As String, valX As String, X1 As String, V As Variant, curX As Currency, dblX As Double
Dim intReturn As Integer
Dim BIA_NET_CMD_Aut As typeAuthorization
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

Dim mMONAPP As String, mMONSTATUS As String
Dim paramService_Start_bat As String
Dim paramService_Stop_bat As String
Dim paramService_Log_bat As String
Dim paramService_Log As String
Dim paramService_Nom As String

'______________________________________________________________________

Dim xYBIAMON7   As typeYBIAMON0
Dim oldYBIAMON7_Service  As typeYBIAMON0, newYBIAMON7_Service  As typeYBIAMON0
Dim oldYBIAMON7_Monitor  As typeYBIAMON0, newYBIAMON7_Monitor  As typeYBIAMON0
Dim blnYBIAMON7_Init As Boolean
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
Private Sub fgSelect_Display_1()
Dim I As Long, x As String
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
cmdSelect_Ok_Caption = "??"

cmdSelect_SQL_1

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub
Private Sub cmdSelect_SQL_1()
Dim V
Dim xSql As String, K As Long
Dim xWhere As String, xAnd As String
On Error GoTo Error_Handler

blnYBIAMON7_Init = False
Set rsSab = Nothing
Call lstErr_Clear(lstErr, cmdContext, "cmdSelect_SQL"): DoEvents

currentAction = "cmdSelect_SQL"
Set rsSab = Nothing

xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIAMON7 where MONAPP = '" & mMONAPP & "' and MONFLUX = 'SERVICE'"
Set rsSab = cnsab.Execute(xSql)
If rsSab.EOF Then
    V = "manque " & xSql
    GoTo Error_MsgBox
End If
V = rsYBIAMON0_GetBuffer(rsSab, oldYBIAMON7_Service)
If Not IsNull(V) Then GoTo Error_MsgBox

If Trim(oldYBIAMON7_Service.MONSTATUS) = "" Then
    mMONSTATUS = "START"
    cmdUpdate_Ok.Caption = mMONSTATUS & " " & mMONAPP
    lblService_MONSTATUS.Caption = "arrêté"
Else
    mMONSTATUS = "STOP"
    cmdUpdate_Ok.Caption = mMONSTATUS & " " & mMONAPP
    lblService_MONSTATUS.Caption = "démarré"
End If
lblService_MONAMJ = oldYBIAMON7_Service.MONUSR & " " & dateImp(oldYBIAMON7_Service.MONAMJ) & " " & timeImp(oldYBIAMON7_Service.MONHMS)



xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIAMON7 where MONAPP = '" & mMONAPP & "' and MONFLUX = 'MONITOR'"
Set rsSab = cnsab.Execute(xSql)
If rsSab.EOF Then
    V = "manque " & xSql
    GoTo Error_MsgBox
End If
V = rsYBIAMON0_GetBuffer(rsSab, oldYBIAMON7_Monitor)
If Not IsNull(V) Then GoTo Error_Handler
lblMonitor_MONSTATUS.Caption = oldYBIAMON7_Monitor.MONSTATUS
lblMonitor_MONAMJ = oldYBIAMON7_Monitor.MONUSR & " " & dateImp(oldYBIAMON7_Monitor.MONAMJ) & " " & timeImp(oldYBIAMON7_Monitor.MONHMS)
  
    
blnYBIAMON7_Init = True
cmdUpdate_Ok.Enabled = True

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Public Sub cmdSendMail_Service()
Dim wSendMail As typeSendMail
Dim bgColor As String
Dim x As String

wSendMail.FromDisplayName = "MONITOR"
wSendMail.RecipientDisplayName = mMONAPP
If Trim(newYBIAMON7_Service.MONSTATUS) = "" Then
    x = " est arrêté."
    bgColor = "<body bgcolor = #FF0000>"
Else
    x = " est démarré."
    bgColor = "<body bgcolor = #A0FFA0>"
End If

wSendMail.Subject = "Le service " & mMONAPP & x
wSendMail.Attachment = ""
wSendMail.Message = bgColor
wSendMail.AsHTML = True

'xxxx DR 30/11/2010 version 1.0.2
'srvSendMail.Monitor wSendMail

End Sub
Public Sub cmdSendMail_Alerte(lFct As String)
Dim wSendMail As typeSendMail
Dim bgColor As String
Dim x As String

wSendMail.FromDisplayName = "ALERTE"
wSendMail.RecipientDisplayName = mMONAPP
    bgColor = "<body bgcolor = #FF0000>"

wSendMail.Subject = "Le service " & mMONAPP & " ne répond pas à la procèdure " & lFct
wSendMail.Attachment = ""
wSendMail.Message = bgColor
wSendMail.AsHTML = True

srvSendMail.Monitor wSendMail

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
Dim I As Integer, x As String
Dim wIndex As Integer
For I = 1 To fgSelect.Rows - 1
    fgSelect.Row = I
    fgSelect.Col = fgSelect_arrIndex
    wIndex = Val(fgSelect.Text)
    Select Case lK
    End Select
    fgSelect.Col = fgSelect_arrIndex - 1
    fgSelect.Text = x
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
cmdSelect_SQL_K = mId$(cboSelect_SQL, 1, 1)
If blnControl Then
    Me.Enabled = False: Me.MousePointer = vbHourglass
    Select Case cmdSelect_SQL_K
        Case "1": mMONAPP = "ADES": lstSelect_Load_1:
    End Select
    Me.Enabled = True: Me.MousePointer = 0
End If
End Sub


Private Sub cmdContext_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
MouseMoveActiveControl_Set cmdContext

End Sub

Private Sub cmdPrint_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
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
blnControl = True
'cboSelect_SQL.ListIndex = 0
End Sub
Public Sub Form_Init()
Call lstErr_Clear(lstErr, cmdContext, "Initialisation ")
DoEvents

SSTab1.Tab = 0

blnControl = False

fgSelect.Visible = False
fgSelect_FormatString = fgSelect.FormatString
cboSelect_SQL.Clear

cboSelect_SQL.AddItem "1 - ADES":
mMONAPP = "ADES"

cmdReset

lblService_MONSTATUS.ForeColor = vbMagenta
lblMonitor_MONSTATUS.ForeColor = vbMagenta


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

Private Sub cmdUpdate_Ok_Click()
Dim V
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> Début du traitement"): DoEvents

If IsNull(cmdUpdate_Control) Then
    Call lstErr_AddItem(lstErr, cmdContext, ">_________Enregistrement des données "): DoEvents
    V = cmdUpdate_Ok_Transaction("MONITOR")
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    If IsNull(V) Then
        cmdUpdate_Ok.Enabled = False
        Unload Me
    Else
        MsgBox V, vbCritical, Me.Name & " : cmdUpdate_Ok"
        Call lstErr_AddItem(lstErr, cmdContext, V): DoEvents
    
    End If
End If
Call lstErr_AddItem(lstErr, cmdContext, "< Fin du Traitement"): DoEvents

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdUpdate_Quit_Click()
Unload Me
End Sub

Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
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
    If fgSelect.Rows > 1 Then
        Me.Enabled = False: Me.MousePointer = vbHourglass
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

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
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
Dim meUnit As typeUnit, x As String
mWindowState = Me.WindowState
If Me.WindowState <> frmUsr_Windowstate Then Me.WindowState = frmUsr_Windowstate

Call BiaPgmAut_Init(mId$(Msg, 1, 12), BIA_NET_CMD_Aut)

blnSetfocus = True
Form_Init
blnAuto = False
blnYBIAMON7_Init = False
'Msg = "@ADES_SERVIC"
'MsgBox Msg
Call lstErr_Clear(lstErr, cmdContext, Msg): DoEvents

Select Case UCase$(Trim(mId$(Msg, 1, 12)))
    Case "@ADES_SERVIC": blnAuto = True:
                        Call param_Init(mMONAPP)
                        Call Auto_Service
    Case Else: blnAuto = False: cmdSelect_SQL_1
                
                    
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





Public Function cmdUpdate_Control()
Dim wLien As Long

cmdUpdate_Control = Null
newYBIAMON7_Monitor = oldYBIAMON7_Monitor
newYBIAMON7_Monitor.MONSTATUS = mMONSTATUS

End Function

Public Function cmdUpdate_Ok_Transaction(lFct As String)
Dim V, x As String, xSql As String
Dim NB As Long
Dim mMsgBox As String
On Error GoTo Error_Handler
'-------------------------------------------------------
App_Debug = "cmdUpdate_Ok_Transaction"
'-------------------------------------------------------
cmdUpdate_Ok_Transaction = Null
'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
Select Case lFct
    Case "SERVICE": V = sqlYBIAMON0_Update(newYBIAMON7_Service, oldYBIAMON7_Service, True)
    Case "MONITOR": V = sqlYBIAMON0_Update(newYBIAMON7_Monitor, oldYBIAMON7_Monitor, True)
End Select
If Not IsNull(V) Then GoTo Error_MsgBox
GoTo Exit_Sub
'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
Exit_Sub:
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")
    End If
    
    cmdUpdate_Ok_Transaction = V
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

End Function


Public Sub Auto_Service()
Dim V
Dim x As String, xSql As String
Me.Enabled = False
currentAction = "Auto_Service"
Call lstErr_AddItem(lstErr, cmdContext, "Auto_service"): DoEvents

Auto_Service_Status

xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIAMON7 where MONAPP = '" & mMONAPP & "' and MONFLUX = 'MONITOR'"
Set rsSab = cnsab.Execute(xSql)
If rsSab.EOF Then
    V = "manque " & xSql
    GoTo Exit_Sub
End If
V = rsYBIAMON0_GetBuffer(rsSab, oldYBIAMON7_Monitor)
If Not IsNull(V) Then GoTo Exit_Sub

If Trim(oldYBIAMON7_Monitor.MONSTATUS) = "START" _
And Trim(oldYBIAMON7_Service.MONSTATUS) = "" Then
    V = Auto_Service_CMD(paramService_Start_bat, "ON")
Else
    If Trim(oldYBIAMON7_Monitor.MONSTATUS) = "STOP" _
    And Trim(oldYBIAMON7_Service.MONSTATUS) = "ON" Then
        V = Auto_Service_CMD(paramService_Stop_bat, "")
    End If
End If

If IsNull(V) Then
    If Trim(oldYBIAMON7_Monitor.MONSTATUS) <> "" Then
        newYBIAMON7_Monitor = oldYBIAMON7_Monitor
        newYBIAMON7_Monitor.MONSTATUS = ""
        Call cmdUpdate_Ok_Transaction("MONITOR")
    End If
End If

Exit_Sub:
    Unload Me
End Sub
Public Sub Auto_Service_Status()
Dim x As String
currentAction = "Auto_Service_Status"

cmdSelect_SQL_1

If Not blnYBIAMON7_Init Then
    MsgBox "? blnYBIAMON7_Init"
    Exit Sub
End If

newYBIAMON7_Service = oldYBIAMON7_Service
newYBIAMON7_Service.MONSTATUS = Auto_Service_Log
If newYBIAMON7_Service.MONSTATUS = "?" Then
    MsgBox "? Auto_Service_Log"
    Exit Sub
End If
If oldYBIAMON7_Service.MONSTATUS <> newYBIAMON7_Service.MONSTATUS Then
    Auto_Service_Update
    'MsgBox "Mail status"
    cmdSendMail_Service
Else
    If oldYBIAMON7_Service.MONAMJ < DSys _
    Or oldYBIAMON7_Service.MONHMS + 10000 < time_Hms Then
        Auto_Service_Update
    End If
End If
End Sub
Public Sub Auto_Service_Update()
Dim x As String
currentAction = "Auto_Service_Update"

V = cmdUpdate_Ok_Transaction("SERVICE")
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
If IsNull(V) Then
    oldYBIAMON7_Service = newYBIAMON7_Service
Else
    MsgBox V, vbCritical, Me.Name & " : cmdUpdate_Ok"
    Call lstErr_AddItem(lstErr, cmdContext, V): DoEvents

End If
End Sub

Public Function Auto_Service_Log() As String
Dim xStatut As String, iRun As Integer
Dim IdShell As Variant, xIn As String
Dim iWait As Integer
Dim blnOk As Boolean, blnLog As Boolean
currentAction = "Auto_Service_Log"
On Error Resume Next  'Error_Handler
Call lstErr_AddItem(lstErr, cmdContext, "Auto_service_Log"): DoEvents

xStatut = "?"

For iRun = 1 To 100
    DoEvents
'_____________________________________
    IdShell = Shell(paramService_Log_bat, 0)
    
    
    DoEvents
    blnOk = False
    blnLog = False
    
    If IdShell > 0 Then
            Sleep 1000
            For iWait = 1 To 50
                DoEvents
                If Dir(paramService_Log) <> "" Then
                    Open paramService_Log For Input As #1
                    blnLog = True
                    xStatut = ""
                    Do Until EOF(1)
                        Line Input #1, xIn
                        xIn = UCase$(xIn)
                       If InStr(xIn, paramService_Nom) > 0 Then blnOk = True: xStatut = "ON": Exit Do
                    Loop
                End If
                If blnLog Then Exit For
                Sleep 1000
            Next iWait
    End If
    If blnLog Then
        Close #1
        msFileSystem.DeleteFile paramService_Log
    End If
'_____________________________________
    If xStatut <> "?" Then Exit For
    Sleep 100 * iRun
Next iRun

Auto_Service_Log = xStatut

Exit Function

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    Call lstErr_AddItem(lstErr, cmdContext, Me.Name & " ? " & currentAction): DoEvents

End Function

Public Function param_Init(lService) As String
Dim xName As String, xMemo As String
Dim x As String, xIn As String
Dim K1 As Integer, K2 As Integer
Dim blnOk As Boolean

paramService_Nom = ""
paramService_Log = ""

x = lService & "_Start"
currentAction = "param_Init : " & x

V = rsElpTable_Read("Server", "Application", x, xName, xMemo)
If Not IsNull(V) Then GoTo Error_MsgBox
paramService_Start_bat = paramServer(xMemo)
Call lstErr_AddItem(lstErr, cmdContext, paramService_Start_bat): DoEvents

'MsgBox "param_init : forçage TEST à supprimer"
'paramService_Start_bat = "c:\temp\procs\demarreAdes.bat"

If Trim(Dir(paramService_Start_bat)) = "" Then
    V = "fichier n'existe pas : " & paramService_Start_bat
    GoTo Error_MsgBox
End If
Open paramService_Start_bat For Input As #1
Do Until EOF(1)
    Line Input #1, xIn
    xIn = UCase$(xIn)
   If InStr(xIn, "NET START") > 0 Then
        K1 = InStr(xIn, Asc34)
        If K1 > 0 Then K2 = InStr(K1 + 1, xIn, Asc34)
        If K2 > 0 Then paramService_Nom = mId$(xIn, K1 + 1, K2 - K1 - 1)
        Exit Do
   End If
Loop
Close #1
If paramService_Nom = "" Then
    V = "préciser le nom du service : " & paramService_Start_bat
    GoTo Error_MsgBox
End If
Call lstErr_AddItem(lstErr, cmdContext, paramService_Nom): DoEvents

'________________________________________________________________________________
blnOk = False
x = lService & "_Stop"
currentAction = "param_Init : " & x

V = rsElpTable_Read("Server", "Application", x, xName, xMemo)
If Not IsNull(V) Then GoTo Error_MsgBox
paramService_Stop_bat = paramServer(xMemo)
Call lstErr_AddItem(lstErr, cmdContext, paramService_Stop_bat): DoEvents

'paramService_Stop_bat = "c:\temp\procs\ArretAdes.bat"

If Trim(Dir(paramService_Stop_bat)) = "" Then
    V = "fichier n'existe pas : " & paramService_Stop_bat
    GoTo Error_MsgBox
End If
Open paramService_Stop_bat For Input As #1
Do Until EOF(1)
    Line Input #1, xIn
    xIn = UCase$(xIn)
   If InStr(xIn, "NET STOP") > 0 Then
        If InStr(xIn, paramService_Nom) > 0 Then blnOk = True: Exit Do
   End If
Loop
Close #1
If Trim(Dir(paramService_Stop_bat)) = "" Then
    V = "nom du service START <> nom du service STOP : " & paramService_Stop_bat
    GoTo Error_MsgBox
End If
'________________________________________________________________________________
blnOk = False
x = lService & "_Log"
currentAction = "param_Init : " & x

V = rsElpTable_Read("Server", "Application", x, xName, xMemo)
If Not IsNull(V) Then GoTo Error_MsgBox
paramService_Log_bat = paramServer(xMemo)
Call lstErr_AddItem(lstErr, cmdContext, paramService_Log_bat): DoEvents

'paramService_Log_bat = "c:\temp\procs\NET_Service_log.bat"

If Trim(Dir(paramService_Log_bat)) = "" Then
    V = "fichier n'existe pas : " & paramService_Log_bat
    GoTo Error_MsgBox
End If
Open paramService_Log_bat For Input As #1
Do Until EOF(1)
    Line Input #1, xIn
    xIn = UCase$(xIn)
    If InStr(xIn, "NET START") > 0 Then
        K1 = InStr(xIn, "> ")
        If K1 > 0 Then
            paramService_Log = mId$(xIn, K1 + 2, Len(xIn) - K1 + 1)
            Exit Do
        End If
    End If
Loop
Close #1

If paramService_Log = "" Then
    V = "préciser le nom du fichier log : " & paramService_Log_bat
    GoTo Error_MsgBox
End If
'___________________________________________________________
Exit Function

Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, Me.Name & " : " & currentAction
    End
End Function



Public Function Auto_Service_CMD(lFct As String, lStatus As String)
Dim IdShell As Variant, x As String
Dim iWait As Integer
Dim blnOk As Boolean, blnLog As Boolean
currentAction = "Auto_Service_CMD"
Call lstErr_AddItem(lstErr, cmdContext, "Auto_service_CMD : " & lFct): DoEvents

Auto_Service_CMD = "?"
x = lFct & " GO"
'MsgBox "SHELL :" & X, vbInformation, currentAction
IdShell = Shell(x, 0)

DoEvents
For iWait = 1 To 20
    Sleep 100
    x = Auto_Service_Log
    If x <> "?" Then
        If x = lStatus Then Auto_Service_CMD = Null: Exit Function
    End If
Next iWait

Call cmdSendMail_Alerte(lFct)

End Function
