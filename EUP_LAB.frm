VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmEUP_LAB 
   AutoRedraw      =   -1  'True
   Caption         =   "EUP_LAB"
   ClientHeight    =   9150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13560
   Icon            =   "EUP_LAB.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9150
   ScaleWidth      =   13560
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   7800
      TabIndex        =   3
      Top             =   0
      Width           =   5175
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8655
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   13530
      _ExtentX        =   23865
      _ExtentY        =   15266
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "SEPA :EUP_LAB"
      TabPicture(0)   =   "EUP_LAB.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraTab0"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "....."
      TabPicture(1)   =   "EUP_LAB.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      Begin VB.Frame fraTab0 
         Height          =   8205
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   13290
         Begin VB.ComboBox cboSelect_SQL 
            Height          =   315
            Left            =   9480
            TabIndex        =   9
            Text            =   "cboSelect_SQL"
            Top             =   120
            Width           =   3615
         End
         Begin VB.Frame fraSelect_Options 
            Height          =   1005
            Left            =   120
            TabIndex        =   7
            Top             =   120
            Width           =   6075
         End
         Begin VB.CommandButton cmdSelect_Ok 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Exécuter la requête"
            Height          =   525
            Left            =   11040
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   600
            Width           =   1095
         End
         Begin MSFlexGridLib.MSFlexGrid fgSelect 
            Height          =   6825
            Left            =   120
            TabIndex        =   8
            Top             =   1200
            Width           =   12840
            _ExtentX        =   22648
            _ExtentY        =   12039
            _Version        =   393216
            Rows            =   1
            Cols            =   22
            FixedCols       =   0
            RowHeightMin    =   450
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
            FormatString    =   $"EUP_LAB.frx":047A
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
      Left            =   13080
      Picture         =   "EUP_LAB.frx":05BF
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   500
   End
   Begin VB.Label libSelect 
      BackColor       =   &H00FFFED9&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   4905
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuContext 
      Caption         =   "mnuContext"
      Visible         =   0   'False
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
      Begin VB.Menu mnuSelect_Print_Liste 
         Caption         =   "Imprimer liste"
      End
      Begin VB.Menu mnuSelect_Print_Détail 
         Caption         =   "Imprimer liste détaillée"
      End
   End
End
Attribute VB_Name = "frmEUP_LAB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------
Option Explicit

Dim mHeight_0 As Integer, mWidth_0 As Integer, mHeight_2 As Integer, mWidth_2 As Integer, mWindowState As Integer
Dim arrTag() As Boolean, arrTagNb As Integer
Dim lastActiveControl_Name  As String, currentActiveControl_Name As String, currentAction As String
Dim MouseMoveActiveControl_Name  As String, MouseMoveActiveControl As typeUsrColor
Dim blnMsgBox_Quit As Boolean, blnAddNew As Boolean, blnGlobalControl As Boolean, blnControl As Boolean
Dim X As String, I As Integer, Msg As String, valX As String, X1 As String, V As Variant, curX As Currency, dblX As Double
Dim intReturn As Integer
Dim EUPLAB_Aut As typeAuthorization
Dim curX1 As Currency, curX2 As Currency
Dim blnAuto As Boolean

Dim fgSelect_FormatString As String, fgSelect_K As Integer
Dim fgSelect_RowDisplay As Integer, fgSelect_RowClick As Integer, fgSelect_ColClick As Integer
Dim fgSelect_ColorClick As Long, fgSelect_ColorDisplay As Long
Dim fgSelect_Sort1 As Integer, fgSelect_Sort2 As Integer
Dim fgSelect_SortAD As Integer, fgSelect_Sort1_Old As Integer
Dim fgSelect_arrIndex As Integer
Dim blnfgSelect_DisplayLine As Boolean

Dim blnTransaction As Boolean
Dim cmdSelect_SQL_K As String
'______________________________________________________________________

Dim wAmjMin As String, wAmjMax As String, wHmsMin As Long, wHmsMax As Long
Dim xSideEUPLAB0 As typeSideEUPLAB0, newSideEUPLAB0 As typeSideEUPLAB0, oldSideEUPLAB0 As typeSideEUPLAB0
Dim arrSideEUPLAB0() As typeSideEUPLAB0, arrSideEUPLAB0_Nb As Long, arrSideEUPLAB0_Max As Long, arrSideEUPLAB0_Index As Long

Dim xYEUPMON0 As typeYEUPMON0, newYEUPMON0 As typeYEUPMON0, oldYEUPMON0 As typeYEUPMON0
Dim arrYEUPMON0() As typeYEUPMON0, arrYEUPMON0_Nb As Long, arrYEUPMON0_Max As Long, arrYEUPMON0_Index As Long

'______________________________________________________________________



Dim cnSQL_Server_BIA As New ADODB.Connection, rsSQL_Server_BIA As New ADODB.Recordset
Dim cnSQL_Server_BIA_VM As New ADODB.Connection, rsSQL_Server_BIA_VM As New ADODB.Recordset
Dim meSideEUPLAB0_Status As typeSideEUPLAB0, oldSideEUPLAB0_Status As typeSideEUPLAB0
Dim paramSafeWatch_Path As String

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
Private Sub fgSelect_Display_1()
Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
fgSelect_Reset

fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
currentAction = "fgselect_Display"
    
For I = 1 To arrYEUPMON0_Nb
         
    xYEUPMON0 = arrYEUPMON0(I)
        fgSelect.Rows = fgSelect.Rows + 1
        fgSelect.Row = fgSelect.Rows - 1
        fgSelect_DisplayLine_1 I
Next I

fgSelect.Visible = True
Call lstErr_AddItem(lstErr, cmdContext, "Opérations : " & arrYEUPMON0_Nb): DoEvents
'If fgSelect.Rows > 1 Then
'    fgSelect_Sort1 = 0: fgSelect_Sort2 = 2: fgSelect_Sort
'End If

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub

Private Sub fgSelect_Display_2()
Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
fgSelect_Reset

fgSelect.Rows = 1
fgSelect.FormatString = "Identification                      |BIC          |Nom                               |> Montant      |Dev |<Scan Status     |<Detection ID          |<Scan Date/Time                |<Violation Summary                           |Rank |<Alert Summary      |<Violation count|Accept count |<Libellé                              |"
currentAction = "fgselect_Display"
    
For I = 1 To arrSideEUPLAB0_Nb
         
    xSideEUPLAB0 = arrSideEUPLAB0(I)
        fgSelect.Rows = fgSelect.Rows + 1
        fgSelect.Row = fgSelect.Rows - 1
        fgSelect_DisplayLine_2 I
Next I

fgSelect.Visible = True
Call lstErr_AddItem(lstErr, cmdContext, "Opérations : " & arrSideEUPLAB0_Nb): DoEvents
'If fgSelect.Rows > 1 Then
'    fgSelect_Sort1 = 0: fgSelect_Sort2 = 2: fgSelect_Sort
'End If

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub


Private Sub arrSideEUPLAB0_SQL(xWhere As String)
Dim V
Dim X As String, xSql As String
On Error GoTo Error_Handler
ReDim arrSideEUPLAB0(101)
arrSideEUPLAB0_Max = 100: arrSideEUPLAB0_Nb = 0

Set rsSQL_Server_BIA = Nothing

xSql = "select * from  " & paramODBC_SideEUPLAB0 & xWhere & " order by EUPLABID"
Set rsSQL_Server_BIA = cnSQL_Server_BIA.Execute(xSql)

Do While Not rsSQL_Server_BIA.EOF
    V = rsSideEUPLAB0_GetBuffer(rsSQL_Server_BIA, xSideEUPLAB0)

     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "frmEUP_LAB.fgselect_Display"
        '' Exit Sub
     Else
         arrSideEUPLAB0_Nb = arrSideEUPLAB0_Nb + 1
         If arrSideEUPLAB0_Nb > arrSideEUPLAB0_Max Then
             arrSideEUPLAB0_Max = arrSideEUPLAB0_Max + 50
             ReDim Preserve arrSideEUPLAB0(arrSideEUPLAB0_Max)
         End If
         
         arrSideEUPLAB0(arrSideEUPLAB0_Nb) = xSideEUPLAB0
    End If
    rsSQL_Server_BIA.MoveNext

Loop

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    Error_Route V
    'If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub
Private Sub arrYEUPMON0_SQL(xWhere As String)
Dim V
Dim X As String, xSql As String
On Error GoTo Error_Handler
ReDim arrYEUPMON0(101)
arrYEUPMON0_Max = 100: arrYEUPMON0_Nb = 0

Set rsSab = Nothing

xSql = "select * from " & paramIBM_Library_SABSPE & ".YEUPMON0 " & xWhere & " order by EUPMONID"
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    V = rsYEUPMON0_GetBuffer(rsSab, xYEUPMON0)

     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "frmEUP_LAB.fgselect_Display"
        '' Exit Sub
     Else
         arrYEUPMON0_Nb = arrYEUPMON0_Nb + 1
         If arrYEUPMON0_Nb > arrYEUPMON0_Max Then
             arrYEUPMON0_Max = arrYEUPMON0_Max + 50
             ReDim Preserve arrYEUPMON0(arrYEUPMON0_Max)
         End If
         
         arrYEUPMON0(arrYEUPMON0_Nb) = xYEUPMON0
    End If
    rsSab.MoveNext

Loop

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    Error_Route V
    'If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Public Sub fgSelect_DisplayLine_1(lIndex As Long)
Dim XX As String
On Error Resume Next
fgSelect.Col = 0: fgSelect.Text = xYEUPMON0.EUPMONID
fgSelect.Col = 1: fgSelect.Text = xYEUPMON0.EUPMONBIC
fgSelect.Col = 2: fgSelect.Text = xYEUPMON0.EUPMONNOM
fgSelect.Col = 3: fgSelect.Text = Format$(xYEUPMON0.EUPMONMON, "### ### ### ###.00")
fgSelect.Col = 4: fgSelect.Text = xYEUPMON0.EUPMONDEV
fgSelect.Col = 5: fgSelect.Text = dateIBM10(xYEUPMON0.EUPMONECH, True)
fgSelect.Col = 6: fgSelect.Text = xYEUPMON0.EUPMONPRI
fgSelect.Col = 7: fgSelect.Text = xYEUPMON0.EUPMONSTA
fgSelect.Col = 8: fgSelect.Text = dateImp10(xYEUPMON0.EUPMONDMO)
fgSelect.Col = 9
XX = xYEUPMON0.EUPMONHMO
XX = CStr(CLng(XX) + 100000000)
fgSelect.Text = mId(XX, 2, 2) & ":" & mId(XX, 4, 2) & ":" & mId(XX, 6, 2)
fgSelect.Col = 10: fgSelect.Text = dateImp10(xYEUPMON0.EUPMONDSW)
fgSelect.Col = 11
XX = xYEUPMON0.EUPMONHSW
XX = CStr(CLng(XX) + 100000000)
fgSelect.Text = mId(XX, 2, 2) & ":" & mId(XX, 4, 2) & ":" & mId(XX, 6, 2)
fgSelect.Col = 12: fgSelect.Text = xYEUPMON0.EUPMONTIC
fgSelect.Col = 13: fgSelect.Text = xYEUPMON0.EUPMONDID
fgSelect.Col = 14: fgSelect.Text = xYEUPMON0.EUPMONLIB
fgSelect.Col = 15: fgSelect.Text = xYEUPMON0.EUPG2AOPE
fgSelect.Col = 16: fgSelect.Text = xYEUPMON0.EUPG2ANUM
fgSelect.Col = 17: fgSelect.Text = xYEUPMON0.EUPG2ACRE
fgSelect.Col = 18: fgSelect.Text = xYEUPMON0.EUPG2ANEC

fgSelect.Col = 19: fgSelect.Text = xYEUPMON0.EUPMONSEQ

fgSelect.Col = fgSelect_arrIndex: fgSelect.Text = lIndex



End Sub

Public Sub fgSelect_DisplayLine_2(lIndex As Long)
On Error Resume Next
fgSelect.Col = 0: fgSelect.Text = xSideEUPLAB0.EUPLABID
fgSelect.Col = 1: fgSelect.Text = xSideEUPLAB0.EUPLABBICE
fgSelect.Col = 2: fgSelect.Text = xSideEUPLAB0.EUPLABNOME
fgSelect.Col = 3: fgSelect.Text = Format$(xSideEUPLAB0.EUPLABMONT, "### ### ### ###.00")
fgSelect.Col = 4: fgSelect.Text = Trim(xSideEUPLAB0.EUPLABDEVI)
fgSelect.Col = 5: fgSelect.Text = xSideEUPLAB0.EUPLABSTAS1
fgSelect.Col = 6: fgSelect.Text = xSideEUPLAB0.EUPLABSTAS2
fgSelect.Col = 7: fgSelect.Text = xSideEUPLAB0.EUPLABSTAS3
fgSelect.Col = 8: fgSelect.Text = xSideEUPLAB0.EUPLABSTAS4
fgSelect.Col = 9: fgSelect.Text = xSideEUPLAB0.EUPLABSTAS5
fgSelect.Col = 10: fgSelect.Text = xSideEUPLAB0.EUPLABSTAS6
fgSelect.Col = 11: fgSelect.Text = xSideEUPLAB0.EUPLABSTAS7
fgSelect.Col = 12: fgSelect.Text = xSideEUPLAB0.EUPLABSTAS8
fgSelect.Col = 13: fgSelect.Text = xSideEUPLAB0.EUPLABLIB
fgSelect.Col = 14: fgSelect.Text = xSideEUPLAB0.EUPLABSTAI
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


Public Sub Msg_Rcv(Msg As String)
'---------------------------------------------------------
mWindowState = Me.WindowState
If Me.WindowState <> frmUsr_Windowstate Then Me.WindowState = frmUsr_Windowstate

Call BiaPgmAut_Init(mId$(Msg, 1, 12), EUPLAB_Aut)
Form_Init


Select Case UCase$(Trim(mId$(Msg, 1, 12)))
    Case "@EUP_LAB": blnAuto = True:
                    cmdSelect_SQL_Auto
''$JPL 2012-12-17 ________________________________________________________________________________
''___________________________________________________________________________________________
''                    'paramIBM_Library_SABSPE = "SAB073USPE"
''                    cmdSelect_SQL_5
''                    Sleep 5000
''                    cmdSelect_SQL_6
''                    Sleep 5000
''                    cmdSelect_SQL_7
''                    ' 18/12/2009 - Modification temps pour exécution lot > 200 opérations
''                    'Sleep 5000
''                    Sleep 20000
''                    cmdSelect_SQL_8
''                    'paramIBM_Library_SABSPE = "SAB073SPE"
''                    Unload Me
''    'Case "@SMS_EUP_LAB": blnAuto = True:
''     '               SMS_EUP_LAB
''     '               Unload Me
''$JPL 2012-12-17 ________________________________________________________________________________
    Case Else: blnAuto = False
                'If Not blnOff_Line Then
                '    MsgBox "en TEST : SAB073U ", vbExclamation
                '    paramIBM_Library_SABSPE = "SAB073USPE"
                'End If

End Select


End Sub


Public Sub cmdSendMail_Alerte(lSubject As String, lMessage As String)
Dim wSendMail As typeSendMail
Dim bgColor As String
Dim X As String

wSendMail.FromDisplayName = "ALERTE"
wSendMail.RecipientDisplayName = "EUP_LAB"
bgColor = "<body bgcolor = #FF0000>"

wSendMail.Subject = lSubject
wSendMail.Attachment = ""
wSendMail.Message = bgColor & lMessage
wSendMail.AsHTML = True

srvSendMail.Monitor wSendMail

End Sub

Public Sub cmdSendMail_Rejet(lK2 As String, lSubject As String, lMessage As String)
Dim wSendMail As typeSendMail
Dim bgColor As String
Dim X As String

wSendMail.FromDisplayName = lK2
wSendMail.RecipientDisplayName = "EUP_LAB"

Select Case lK2
    Case "SEPA_LAB_FIN": bgColor = "<body bgcolor = #A0FFA0>"
    Case Else: bgColor = "<body bgcolor = #FF80FF>"
End Select
wSendMail.Subject = lSubject
wSendMail.Attachment = ""
wSendMail.Message = bgColor & lMessage
wSendMail.AsHTML = True

srvSendMail.Monitor wSendMail

End Sub

Public Sub Form_Init()
Me.Enabled = False
Me.MousePointer = vbHourglass

Call lstErr_Clear(lstErr, cmdPrint, "Initialisation ")
DoEvents

If Not IsNull(param_Init) Then
    If Not blnAuto Then MsgBox "paramétrage inconsistant", vbCritical, "frmEUPLAB.param_init"
    Unload Me
Else
    lstErr.Clear
End If


blnControl = False
fgSelect_FormatString = fgSelect.FormatString

fgSelect.Enabled = True
cmdReset

Me.Enabled = True
Me.MousePointer = 0
End Sub


Private Sub cboSelect_SQL_Click()
cmdSelect_SQL_K = mId$(cboSelect_SQL, 1, 1)

fraSelect_Options.Enabled = True
cmdSelect_Ok_Click
End Sub



'---------------------------------------------------------
Public Sub cmdReset()
'---------------------------------------------------------
blnControl = False
usrColor_Set
cmdContext.Caption = constcmdRechercher: blnMsgBox_Quit = False
arrTag_Set False
lstErr.Visible = False
currentAction = ""
fraSelect_Options.Enabled = True
'cmdSelect_Ok_Click



blnControl = True



End Sub


Public Function param_Init()
Dim xName As String, xMemo As String
param_Init = Null
Call lstErr_Clear(lstErr, cmdContext, ". EUP_LAB_Import cbo"): DoEvents

fgSelect.Visible = False
V = rsElpTable_Read("Server", "Application", "EUP_LAB_Proc", xName, xMemo)

paramSafeWatch_Path = xMemo '"C:\Program Files\SafeWatch\DBConnector_Console\"

cboSelect_SQL.Clear
cboSelect_SQL.AddItem "1 - opérations en attente (SAB)"
cboSelect_SQL.AddItem "2 - opérations en attente (SafeWatch)"
If EUPLAB_Aut.Xspécial Then
    cboSelect_SQL.AddItem "5 - LAB : Extraction (EUP => Side)"
    cboSelect_SQL.AddItem "6 - LAB : SideScan "
    cboSelect_SQL.AddItem "7 - LAB : SideMonitor "
    cboSelect_SQL.AddItem "8 - LAB : Synchronisation (Side => EUP)"
    cboSelect_SQL.AddItem "X - Test JPL"
End If

cboSelect_SQL.ListIndex = 0


Me.Enabled = True: Me.MousePointer = 0

End Function





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



Private Sub cmdContext_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Select Case cmdContext.Caption
    Case Is = constcmdRechercher: Me.PopupMenu mnuContext, vbPopupMenuLeftButton
    Case Is = constcmdAbandonner: cmdContext_Quit
End Select

End Sub

Private Sub cmdPrint_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Select Case SSTab1.Tab
    Case 0:
            If fgSelect.Rows > 1 Then cmdPrint_Ok
                
    Case 1:
End Select
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdSelect_SQL_1()
Dim V
Dim X As String
Dim xWhere As String, xAnd As String
Dim wAmj7 As Long
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_1"
xWhere = "where EUPMONSTA >= '0' and EUPMONSTA <= '9'  and EUPMONMON > 0.01"

Call arrYEUPMON0_SQL(xWhere)

fgSelect_Display_1

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub cmdSelect_SQL_5()
Dim V
Dim X As String
Dim xSql As String
Dim wAmj7 As Long
On Error GoTo Error_Handler

App_Debug = "cmdSelect_SQL_5"

currentAction = "cmdSelect_SQL_5"
xSql = "where EUPMONSTA = '0' and EUPMONMON > 0.01"

Call arrYEUPMON0_SQL(xSql)
fgSelect_Display_1

If arrYEUPMON0_Nb > 0 Then
    For I = 1 To arrYEUPMON0_Nb
             
        oldYEUPMON0 = arrYEUPMON0(I)
        newYEUPMON0 = oldYEUPMON0
        newYEUPMON0.EUPMONSTA = "1"
        
        rsSideEUPLAB0_Init newSideEUPLAB0
        newSideEUPLAB0.EUPLABID = oldYEUPMON0.EUPMONID
        newSideEUPLAB0.EUPLABBICE = oldYEUPMON0.EUPMONBIC
        newSideEUPLAB0.EUPLABNOME = oldYEUPMON0.EUPMONNOM
        newSideEUPLAB0.EUPLABLIB = oldYEUPMON0.EUPMONLIB
        newSideEUPLAB0.EUPLABMONT = oldYEUPMON0.EUPMONMON
        newSideEUPLAB0.EUPLABDEVI = oldYEUPMON0.EUPMONDEV
        newSideEUPLAB0.EUPLABSTAS1 = "0"
        X = Trim(oldYEUPMON0.EUPMONBFI) & " " & Trim(oldYEUPMON0.EUPMONBF2)
        newSideEUPLAB0.EUPLABSTAS2 = Left(X, 50) ' Bénéficiaire
        xSql = "select * from  " & paramODBC_SideEUPLAB0 & " where EUPLABID = '" & Trim(newSideEUPLAB0.EUPLABID) & "'"
        Set rsSQL_Server_BIA = cnSQL_Server_BIA.Execute(xSql)
        If Not rsSQL_Server_BIA.EOF Then
            V = "l'opération " & newSideEUPLAB0.EUPLABID & " existe déjà dans la table SQL_Server_BIA."
            Error_Route V
        Else
            cmdSelect_SQL_5_Transaction
        End If
    Next I
End If
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    Error_Route V

End Sub

Public Function cmdSelect_SQL_5_Transaction()
Dim V, X As String, xSql As String
Dim Nb As Long
Dim mMsgBox As String
On Error GoTo Error_Handler
'-------------------------------------------------------
App_Debug = "cmdSelect_SQL_5_Transaction"
'-------------------------------------------------------
cmdSelect_SQL_5_Transaction = Null

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
V = sqlYEUPMON0_Update(newYEUPMON0, oldYEUPMON0)
If Not IsNull(V) Then GoTo Error_MsgBox

V = sqlSideEUPLAB0_Insert(newSideEUPLAB0, cnSQL_Server_BIA)
If Not IsNull(V) Then GoTo Error_MsgBox
GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    Error_Route V
Exit_sub:
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")
    End If
    
    cmdSelect_SQL_5_Transaction = V
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


End Function
Public Function cmdSelect_SQL_8_Transaction(blnSW_Delete As Boolean, blnYEUPMON_Update As Boolean)
Dim V, X As String, xSql As String
Dim Nb As Long
Dim mMsgBox As String
On Error GoTo Error_Handler
'-------------------------------------------------------
App_Debug = "cmdSelect_SQL_8_Transaction"
'-------------------------------------------------------
cmdSelect_SQL_8_Transaction = Null

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
If blnSW_Delete Then
    V = sqlSideEUPLAB0_Delete(oldSideEUPLAB0, cnSQL_Server_BIA)
    If Not IsNull(V) Then GoTo Error_MsgBox
End If

If blnYEUPMON_Update Then
    V = sqlYEUPMON0_Update(newYEUPMON0, oldYEUPMON0)
    If Not IsNull(V) Then GoTo Error_MsgBox
End If

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    Error_Route V
Exit_sub:
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")
    End If
    
    cmdSelect_SQL_8_Transaction = V
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


End Function


Private Sub cmdSelect_SQL_2()
Dim V
Dim X As String
Dim xWhere As String, xAnd As String
Dim wAmj7 As Long
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_2"
xWhere = ""

Call arrSideEUPLAB0_SQL(xWhere)

fgSelect_Display_2

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub
Private Sub cmdSelect_SQL_8()
Dim V
Dim X As String, xMessage As String
Dim xSql As String
Dim blnSW_Delete As Boolean, blnYEUPMON_Update As Boolean
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_8"
xSql = " where EUPLABSTAS1 <> '0'"

Call arrSideEUPLAB0_SQL(xSql)

fgSelect_Display_2
If arrSideEUPLAB0_Nb > 0 Then
    For I = 1 To arrSideEUPLAB0_Nb
        blnSW_Delete = False
        blnYEUPMON_Update = False
        
        oldSideEUPLAB0 = arrSideEUPLAB0(I)
        
        xSql = "select * from " & paramIBM_Library_SABSPE & ".YEUPMON0 where EUPMONID = '" & Trim(oldSideEUPLAB0.EUPLABID) & "'"
        Set rsSab = cnsab.Execute(xSql)
        
        If rsSab.EOF Then
            blnSW_Delete = True
        Else
            
            V = rsYEUPMON0_GetBuffer(rsSab, oldYEUPMON0)
            If IsNull(V) Then
               newYEUPMON0 = oldYEUPMON0
               Select Case oldYEUPMON0.EUPMONSTA
                    Case "1", "2": 'statut dans MON0
                        Select Case Trim(oldSideEUPLAB0.EUPLABSTAS1)
                            Case "Scanned", "Accepted": blnYEUPMON_Update = True: blnSW_Delete = True: newYEUPMON0.EUPMONSTA = "V"
                            Case "Real Violation": blnYEUPMON_Update = True
                                                   blnSW_Delete = True
                                                   newYEUPMON0.EUPMONSTA = "R"
                                                   newYEUPMON0.EUPMONTIC = Trim(oldSideEUPLAB0.EUPLABSTAS8)
                                                   newYEUPMON0.EUPMONDID = Trim(oldSideEUPLAB0.EUPLABSTAS1)
                                                   X = "LAB : rejet de l'opération " & oldSideEUPLAB0.EUPLABID
                                                   xMessage = "Statut : " & oldSideEUPLAB0.EUPLABSTAS1 & "<br/>" _
                                                            & "Ticket : " & oldSideEUPLAB0.EUPLABSTAS8 & "<br/>" _
                                                            & "Date   : " & oldSideEUPLAB0.EUPLABSTAS3 & "<br/>" _
                                                            & "Motif  : " & oldSideEUPLAB0.EUPLABSTAS4
                                                    Call cmdSendMail_Rejet("Rejet", X, xMessage)
                            Case "False Positive", "External": blnYEUPMON_Update = True: blnSW_Delete = True: newYEUPMON0.EUPMONSTA = "V"
                            Case "Reported", "New", "Investigating", "Dont Know":
                                                    If oldYEUPMON0.EUPMONSTA = "1" Then
                                                        blnYEUPMON_Update = True
                                                        newYEUPMON0.EUPMONSTA = "2"
                                                        newYEUPMON0.EUPMONTIC = Trim(oldSideEUPLAB0.EUPLABSTAS8)
                                                        newYEUPMON0.EUPMONDID = Trim(oldSideEUPLAB0.EUPLABSTAS1)
                                                    End If

                           Case Else: Call Error_Route("? cmdSelect_Sql_8 : " & Trim(oldSideEUPLAB0.EUPLABSTAS1))
                        End Select
                    Case "A", "$", "V", "R": blnSW_Delete = True
               End Select
            End If

        End If
        If blnSW_Delete Or blnYEUPMON_Update Then Call cmdSelect_SQL_8_Transaction(blnSW_Delete, blnYEUPMON_Update)
    Next I
End If

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    Error_Route V

End Sub

Public Function cmdSelect_SQL_6()
Dim xSql As String, Nb As Long
On Error Resume Next

'xSql = "select count(*) as Tally  from " & paramIBM_Library_SABSPE & ".YEUPMON0  where EUPMONSTA in ('0','1','2')"
xSql = "select count(*) as Tally  from " & paramIBM_Library_SABSPE & ".YEUPMON0  where EUPMONSTA in ('0','1','2') and EUPMONMON > 0.01"
Set rsSab = cnsab.Execute(xSql)
Nb = rsSab("Tally")
If Nb > 0 Then
    'Call Shell_Exe(paramSafeWatch_Path & "SideScan_Shell.bat")
End If
End Function


Public Function cmdSelect_SQL_7()
'=============== La fonction de monitoring (DBMonitor.bat) est lancée directement sur le serveur ENAPP_VM, par une tâche système, toutes les 5 minutes ====
Exit Function
'=====================================================================================================================
Dim xSql As String, Nb As Long
On Error Resume Next
xSql = "select count(*) as Tally  from " & paramIBM_Library_SABSPE & ".YEUPMON0  where EUPMONSTA in ('0','1','2')"
Set rsSab = cnsab.Execute(xSql)
Nb = rsSab("Tally")
If Nb > 0 Then
   Call Shell_Exe(paramSafeWatch_Path & "SideMonitor_Shell.bat")
End If

End Function

Private Sub cmdSelect_Ok_Click()
Dim blnOk As Boolean, Nb As Long

blnOk = fraSelect_Options.Enabled
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> EUP_LAB_cmdSelect_Ok ........"): DoEvents

fgSelect.Clear
If blnOk Then
    cmdSelect_Ok.Caption = "Options"
    cmdSelect_Ok.BackColor = &HFFFFFA   '&HC0FFFF
    fraSelect_Options.BackColor = &H8000000F
    Call usrColor_Container(fraSelect_Options, fraSelect_Options.BackColor)
    fraSelect_Options.Enabled = False
    Select Case cmdSelect_SQL_K
        Case "1":    cmdSelect_SQL_1
        Case "2":    cmdSelect_SQL_2
        Case "5":    cmdSelect_SQL_5
        Case "6":    cmdSelect_SQL_6
        Case "7":    cmdSelect_SQL_7
        Case "8":    cmdSelect_SQL_8
        Case "X": cmdSelect_SQL_Auto
    End Select
Else
    cmdSelect_Ok.Caption = constcmdRechercher
    cmdSelect_Ok.BackColor = &HC0FFC0
    fraSelect_Options.BackColor = &HFFFFFA    '&HC0FFFF
    Call usrColor_Container(fraSelect_Options, fraSelect_Options.BackColor)
    fraSelect_Options.Enabled = True
End If
Call lstErr_AddItem(lstErr, cmdContext, "< EUP_LAB_cmdSelect_Ok"): DoEvents
Me.Enabled = True: Me.MousePointer = 0


End Sub


Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim K As Long
On Error Resume Next
If y <= fgSelect.RowHeightMin Then
    Select Case fgSelect.Col
        Case 0: fgSelect_Sort1 = 0: fgSelect_Sort2 = 2: fgSelect_Sort
        Case 1:  fgSelect_Sort1 = 1: fgSelect_Sort2 = 2: fgSelect_Sort
        Case 2: fgSelect_Sort1 = 2: fgSelect_Sort2 = 2: fgSelect_Sort
        Case 3: fgSelect_Sort1 = 3: fgSelect_Sort2 = 3: fgSelect_Sort
        Case 6: fgSelect_Sort1 = 6: fgSelect_Sort2 = 6: fgSelect_Sort
        Case 7: fgSelect_Sort1 = 7: fgSelect_Sort2 = 7: fgSelect_Sort
        Case 8: fgSelect_Sort1 = 8: fgSelect_Sort2 = 8: fgSelect_Sort
       Case fgSelect_arrIndex:  fgSelect_SortX fgSelect_arrIndex
    End Select
Else
    If fgSelect.Rows > 1 Then
        Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
        fgSelect.Col = fgSelect_arrIndex:  arrSideEUPLAB0_Index = CLng(fgSelect.Text)
        fgSelect.LeftCol = 0
        xSideEUPLAB0 = arrSideEUPLAB0(arrSideEUPLAB0_Index)
        

   End If
End If
fgSelect.LeftCol = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
'If blnControl Then
    cnSQL_Server_BIA.Close
    Set cnSQL_Server_BIA = Nothing
'End If
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
If SSTab1.Tab = 0 Then
        Unload Me
    Exit Sub
Else
    SSTab1.Tab = SSTab1.Tab - 1
End If

End Sub

Public Sub cmdContext_Return()
On Error Resume Next
If SSTab1.Tab = 0 Then
    fgSelect.Row = fgSelect.TopRow
    fgSelect.Col = fgSelect_arrIndex:
Else
    SendKeys "{TAB}"
End If
End Sub

'---------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------
Dim V
Dim xName  As String, xMemo As String
On Error GoTo Error_Handler

mHeight_0 = Me.Height: mWidth_0 = Me.Width: mHeight_2 = 0: mWidth_2 = 0: mWindowState = Me.WindowState
Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)
blnControl = False

cnSQL_Server_BIA.Open paramODBC_DSN_SQL_Server_BIA

Exit Sub

Error_Handler:

blnControl = False
If Not blnAuto Then MsgBox Error
End Sub





Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
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
'Call txt_GotFocus(txt)
'Call txt_LostFocus(txt)

End Sub


Private Sub mnuSelect_Print_Détail_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
cmdPrint_Ok '"D "
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuSelect_Print_Liste_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
cmdPrint_Ok '"L "
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error Resume Next
If SSTab1.Tab = 0 Then cmdSelect_Ok.SetFocus

End Sub














Private Sub SSTab1_GotFocus()
Select Case SSTab1.Tab
    Case 0: fgSelect.LeftCol = 0
End Select
End Sub


Public Sub cmdPrint_Ok()
Dim iRow As Integer, K As Integer, I As Integer
Dim blnOk As Boolean

fgSelect.Visible = False
Call lstErr_Clear(Me.lstErr, Me.cmdContext, "Impression Etat : " & fgSelect.Rows - 1)

fgSelect.Visible = True
Me.Show
End Sub








Public Sub Error_Route(V)

If blnAuto Then
    Call cmdSendMail_Alerte(Me.Name & " ~ " & App_Debug, CStr(V))
Else
    MsgBox V, vbCritical, Me.Name & " ~ " & App_Debug
End If

End Sub

Public Sub cmdSelect_SQL_Auto()
Static mSta0_Nb As Long, mSta0_Loop As Integer, mSta0_Nb_In As Integer
Static mStaNOK_Nb As Long, mStaNOK_Loop As Integer, blnSEPA_URGENT As Boolean

Dim wSta0_Nb As Long, wStaNOK_Nb As Long, wStaOK_Nb As Long
Dim xSql As String

'________________________________________________________________________________________________________
xSql = "select EUPLABSTAS1 , count(*) from  " & paramODBC_SideEUPLAB0 & " group by EUPLABSTAS1"
Set rsSQL_Server_BIA = cnSQL_Server_BIA.Execute(xSql)
Do While Not rsSQL_Server_BIA.EOF
    Select Case rsSQL_Server_BIA(0)
        Case "0": wSta0_Nb = rsSQL_Server_BIA(1)
        Case "Scanned", "Accepted", "Real Violation", "False Positive", "External": wStaOK_Nb = wStaOK_Nb + rsSQL_Server_BIA(1)
        Case Else: wStaNOK_Nb = wStaNOK_Nb + rsSQL_Server_BIA(1)
    End Select
    
    rsSQL_Server_BIA.MoveNext
Loop

If mSta0_Nb_In > 0 Then
    If wSta0_Nb = 0 And wStaOK_Nb = 0 And wStaNOK_Nb = 0 Then
        Call cmdSendMail_Rejet("SEPA_LAB_FIN", "SEPA-Safewatch : contrôle terminé", mSta0_Nb_In & " opérations traitées")
        mSta0_Nb_In = 0
        blnSEPA_URGENT = False
    End If
End If

'________________________________________________________________________________________________________

If wStaOK_Nb > 0 Then
    If paramEnvironnement = constProduction Then Call cmdSelect_SQL_8
'===========================
    Sleep 5000
End If

If wSta0_Nb <> 0 Then
    If mSta0_Nb = wSta0_Nb Then
        If mSta0_Loop = 5 Then
            '06/01/2020 DENIS ROSILLETTE suppression de l'alerte
            'Call cmdSendMail_Alerte("SEPA-Safewatch : DBScan inactif (sur ENAPP_VM)?", "l'application Safewatch ne semble pas répondre : 5 ème tentative du traitement sans résultats")
            
        Else
            If mSta0_Loop < 15 Then
                mSta0_Loop = mSta0_Loop + 1
            Else
                mSta0_Loop = 0
            End If
        End If
    Else
        mSta0_Loop = 0
        mSta0_Nb = wSta0_Nb
    End If
    If paramEnvironnement = constProduction Then Call cmdSelect_SQL_6
'===========================
Else
    mSta0_Nb = 0: mSta0_Loop = 0
        
    If wStaNOK_Nb <> 0 Then
        If mStaNOK_Nb = wStaNOK_Nb Then
            If mStaNOK_Loop = 1 Then
                If Not blnSEPA_URGENT Then
                    blnSEPA_URGENT = True
                    Call cmdSendMail_Rejet("SEPA_URGENT", "SEPA-Safewatch : en attente", "Il y a " & mStaNOK_Nb & " opérations SEPA en attente de décision / " & mSta0_Nb_In & " traitées")
                End If
            Else
                If mStaNOK_Loop < 15 Then
                    mStaNOK_Loop = mStaNOK_Loop + 1
                Else
                    mSta0_Loop = 0
                End If
           End If
        Else
            mStaNOK_Loop = 2
            mStaNOK_Nb = wStaNOK_Nb
        End If
        If paramEnvironnement = constProduction Then Call cmdSelect_SQL_7
'===========================
    Else
        mStaNOK_Nb = 0: mStaNOK_Loop = 0
    End If
End If
'________________________________________________________________________________________________________

xSql = "select count(*) from " & paramIBM_Library_SABSPE & ".YEUPMON0 where EUPMONSTA = '0' and EUPMONMON > 0.01"
Set rsSab = cnsab.Execute(xSql)

If Not rsSab.EOF Then
    If rsSab(0) > 0 Then
        mSta0_Nb = 0: mSta0_Loop = 0
        mSta0_Nb_In = mSta0_Nb_In + rsSab(0)
        Sleep 5000
        If paramEnvironnement = constProduction Then
            Call cmdSelect_SQL_5
            '18/12/2018 DR flag ON dans SIDE2014\flagSepa\flag
            xSql = "update dbo.flagSepa set flag='ON'"
            Call FEU_ROUGE
            Call cnSQL_Server_BIA.Execute(xSql)
            Call FEU_VERT
        End If
'===========================
    Else
        '18/12/2018 DR flag OFF dans SIDE2014\flagSepa\flag
        xSql = "update dbo.flagSepa set flag='OFF'"
        Call FEU_ROUGE
        Call cnSQL_Server_BIA.Execute(xSql)
        Call FEU_VERT
    End If
Else
    '18/12/2018 DR flag OFF dans SIDE2014\flagSepa\flag
    xSql = "update dbo.flagSepa set flag='OFF'"
    Call FEU_ROUGE
    Call cnSQL_Server_BIA.Execute(xSql)
    Call FEU_VERT
End If
End Sub
