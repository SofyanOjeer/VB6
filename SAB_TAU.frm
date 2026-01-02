VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmSAB_TAU 
   AutoRedraw      =   -1  'True
   Caption         =   "SAB_TAU: Interfaces"
   ClientHeight    =   9315
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13575
   Icon            =   "SAB_TAU.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9315
   ScaleWidth      =   13575
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6120
      TabIndex        =   5
      Top             =   0
      Width           =   6900
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8688
      Left            =   -48
      TabIndex        =   3
      Top             =   492
      Width           =   13500
      _ExtentX        =   23813
      _ExtentY        =   15319
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Interface REUTERS"
      TabPicture(0)   =   "SAB_TAU.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraTab0"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "......."
      TabPicture(1)   =   "SAB_TAU.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      Begin VB.Frame fraTab0 
         Height          =   8232
         Left            =   135
         TabIndex        =   4
         Top             =   72
         Width           =   13296
         Begin VB.CommandButton cmdOptions_SAB 
            BackColor       =   &H0000FFFF&
            Caption         =   "SAB"
            Height          =   612
            Left            =   1920
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   960
            Width           =   1176
         End
         Begin VB.CommandButton cmdOptions 
            BackColor       =   &H0000FFFF&
            Caption         =   "Archive Bloomberg"
            Height          =   612
            Left            =   1920
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   240
            Width           =   1176
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
            Height          =   7932
            Left            =   8040
            TabIndex        =   10
            Top             =   2520
            Visible         =   0   'False
            Width           =   3795
            Begin VB.FileListBox filDoc 
               ForeColor       =   &H00008000&
               Height          =   6720
               Left            =   240
               TabIndex        =   11
               Top             =   360
               Visible         =   0   'False
               Width           =   3405
            End
            Begin VB.Label lblContextOptions 
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
               Height          =   612
               Left            =   240
               TabIndex        =   12
               Top             =   7200
               Width           =   3360
            End
         End
         Begin VB.CommandButton cmdReuters_Ok 
            BackColor       =   &H00C0C0FF&
            Caption         =   "Ajouter les taux et les cours ==> SAB"
            Height          =   660
            Left            =   11520
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   960
            Visible         =   0   'False
            Width           =   1416
         End
         Begin VB.Frame fraSelect 
            Height          =   636
            Left            =   3240
            TabIndex        =   6
            Top             =   960
            Visible         =   0   'False
            Width           =   8172
            Begin VB.CommandButton cmdSelect 
               BackColor       =   &H00C0FFC0&
               Caption         =   "lecture fichier Bloomberg"
               Height          =   405
               Left            =   6000
               Style           =   1  'Graphical
               TabIndex        =   7
               Top             =   120
               Width           =   2025
            End
            Begin VB.TextBox txtSelect 
               Height          =   285
               Left            =   75
               TabIndex        =   0
               Top             =   240
               Width           =   5772
            End
         End
         Begin MSFlexGridLib.MSFlexGrid fgSelect 
            Height          =   6468
            Left            =   132
            TabIndex        =   8
            Top             =   1680
            Width           =   12552
            _ExtentX        =   22146
            _ExtentY        =   11404
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
            FormatString    =   $"SAB_TAU.frx":0342
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
         Begin MSComCtl2.DTPicker txtSelect_AMJ 
            Height          =   300
            Left            =   240
            TabIndex        =   15
            Top             =   720
            Width           =   1332
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            CalendarBackColor=   16777215
            CalendarForeColor=   0
            CalendarTitleBackColor=   8421504
            CalendarTitleForeColor=   16777215
            CalendarTrailingForeColor=   12632256
            CustomFormat    =   "dd  MM yyy"
            Format          =   8454147
            CurrentDate     =   38699.44875
            MaxDate         =   401768
            MinDate         =   36526.4425347222
         End
         Begin VB.Label lblSelect 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Fichier à importer"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   732
            Left            =   4080
            TabIndex        =   16
            Top             =   240
            Width           =   5832
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
      Left            =   13080
      Picture         =   "SAB_TAU.frx":03CE
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
   Begin VB.Menu mnufgSelect 
      Caption         =   "mnufgSelect"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmSAB_TAU"
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
Dim SAB_TAU_Aut As typeAuthorization
Dim blnAuto As Boolean, blnError As Boolean

Dim fgSelect_FormatString As String, fgSelect_K As Integer
Dim fgSelect_RowDisplay As Integer, fgSelect_RowClick As Integer, fgSelect_ColClick As Integer
Dim fgSelect_ColorClick As Long, fgSelect_ColorDisplay As Long
Dim fgSelect_Sort1 As Integer, fgSelect_Sort2 As Integer
Dim fgSelect_SortAD As Integer, fgSelect_Sort1_Old As Integer
Dim fgSelect_arrIndex As Integer
Dim blnfgSelect_DisplayLine As Boolean


Dim paramSAB_TAU_Import As String
Dim paramSAB_TAU_Archive As String

Dim meYBASTXX0 As typeYBASTXX0, xYBASTXX0 As typeYBASTXX0
Dim arrYBASTAU0_Key() As String * 16, arrYBASTAU0_Key_Nb As Long, mYBASTAU0_Key As String * 16

Dim mearrYBASTXX0() As typeYBASTXX0
Public mearrYBASTXX0_NB As Integer
Dim blncmdReuters_Ok_Visible As Boolean

Dim newYBASTXX0 As typeYBASTXX0
Dim wAMJMin As String


Dim xYBIAMON0 As typeYBIAMON0, newYBIAMON0 As typeYBIAMON0
Dim DSYS_IBM As Long


Public Sub cmdAuto()
If blnCONTROL_FIX Or blnCONTROL_TAU Then ZBASTAB0_Control

If Not blnUPDATE_TAU And Not blnALERTE_TAU Then
    If time_Hms > 113000 Then cmdSendMail_ALERTE_TAU
End If
If Not blnUPDATE_FIX And Not blnALERTE_FIX Then
    If time_Hms > 160000 Then cmdSendMail_ALERTE_FIX
End If

cmdSelect_Click
If Not blnError Then
    cmdReuters_Ok_Click
'2002.10.22 JPL Shell_MsgBox "# Reuters => SAB : mise à jour des taux & cours terminée # " & Time, vbInformation, Me.Caption, True
End If

Unload Me
End Sub

Private Sub fgSelect_Display()

SSTab1.Tab = 0

fgSelect_Reset

fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString


End Sub

Public Sub fgSelect_DisplayLine(lOrigine As String, lId As String, lText As String)
On Error Resume Next
fgSelect.Rows = fgSelect.Rows + 1
fgSelect.Row = fgSelect.Rows - 1
'fgSelect.Col = 0: fgSelect.Text = lOrigine
'fgSelect.Col = 1: fgSelect.Text = lId
'fgSelect.Col = 2: fgSelect.Text = lText
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
Dim wFct As String
mWindowState = Me.WindowState
If Me.WindowState <> frmUsr_Windowstate Then Me.WindowState = frmUsr_Windowstate

wFct = UCase$(Trim(Mid$(Msg, 1, 12)))
Call BiaPgmAut_Init(wFct, SAB_TAU_Aut)

'blnSetfocus = True
Form_Init

Select Case wFct
    Case "@SAB_TAUX":     blnAuto = True: cmdAuto
    Case Else: blnAuto = False
End Select

End Sub


Public Sub Form_Init()
Dim V
Me.Enabled = False
Call lstErr_Clear(lstErr, cmdPrint, "Initialisation ")
DoEvents
fraContextOptions.Visible = False
fraContextOptions.Top = 200
fraContextOptions.Left = 9360
fraSelect.Visible = SAB_TAU_Aut.Xspécial
cmdReuters_Ok.Visible = False

lstErr.Visible = True
If Not IsNull(param_Init) Then
    MsgBox "paramétrage inconsistant", vbCritical, "frmSAB_TAU.param_init"
    fraTab0.Enabled = False
End If

    blnControl = False
    fgSelect_FormatString = fgSelect.FormatString
    fgSelect.Enabled = True
    cmdReset
    DSYS_IBM = Val(DSys) - 19000000
    Call DTPicker_Set(txtSelect_AMJ, DSys)

    If Not blnSAB_TAUX_Auto Then
        blnSAB_TAUX_Auto = True
         blnEURJ1M = False: blnEURUSD = False
         sabEURJ1M = 0: sabEURUSD = 0
         newEURJ1M = 0: newEURUSD = 0
         Call ZBASTAB0_Load(DSys, True)
        Call paramInit_Auto
   End If
   
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
txtSelect = paramSAB_TAU_Import
txtSelect.Enabled = SAB_TAU_Aut.Xspécial
lblSelect.BackColor = vbYellow 'warnUsrColor
cmdReuters_Ok.Visible = False
blncmdReuters_Ok_Visible = True
mearrYBASTXX0_NB = 0: ReDim mearrYBASTXX0(1)


blnControl = True

End Sub



Public Function param_Init()
Dim K As Integer, K1 As Integer
Dim X As String, xName As String, xMemo As String

Dim V
App_Debug = "frmSAb_TAU.param_Init"
param_Init = Null
On Error GoTo Error_Handler

If SAB_TAU_Aut.Xspécial Then Call lstErr_Clear(Me.lstErr, Me.cmdContext, "BIA.mdb : table : SAB_TAU")



V = rsElpTable_Read("SAB_TAU", "Import", "Reuters", xName, xMemo)
If Not IsNull(V) Then GoTo Error_Handler
paramSAB_TAU_Import = paramServer(xMemo)
If SAB_TAU_Aut.Xspécial Then Call lstErr_AddItem(Me.lstErr, Me.cmdContext, "Reuters : " & paramSAB_TAU_Import)


V = rsElpTable_Read("SAB_TAU", "Archive", "Reuters", xName, xMemo)
If Not IsNull(V) Then GoTo Error_Handler
paramSAB_TAU_Archive = paramServer(xMemo)
If SAB_TAU_Aut.Xspécial Then Call lstErr_AddItem(Me.lstErr, Me.cmdContext, "Archive : " & paramSAB_TAU_Archive)

Exit Function


'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
    param_Init = V

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

Private Sub cmdContext_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Select Case cmdContext.Caption
    Case Is = constcmdRechercher: Me.PopupMenu mnuContext, vbPopupMenuLeftButton
    Case Is = constcmdAbandonner: cmdContext_Quit
End Select

End Sub

Private Sub cmdOptions_Click()
mnuContextOptions_Click
End Sub

Private Sub cmdOptions_SAB_Click()
Call DTPicker_Control(txtSelect_AMJ, wAMJMin)
ZBASTAB0_Load wAMJMin, False
End Sub

Private Sub cmdReuters_Ok_Click()
Dim K As Integer
Dim X1 As String, X2 As String
Dim wIndex As Long, xSQL As String

On Error GoTo Error_Handler
'-------------------------------------------------------
App_Debug = "cmdReuters_Ok_Transaction"
'-------------------------------------------------------

cmdReuters_Ok.Visible = False
blncmdReuters_Ok_Visible = False

Me.Enabled = False: Me.MousePointer = vbHourglass

Call lstErr_Clear(lstErr, cmdContext, "cmdReuters_Ok : début" & Time): DoEvents

V = file_Archive(Trim(txtSelect), paramSAB_TAU_Archive)
Call lstErr_AddItem(lstErr, cmdContext, V)

Call lstErr_AddItem(lstErr, cmdContext, ".....")
'________________________________________________________________________________

fgSelect.Visible = False

newYBASTXX0.BASTXXUAMJ = DSys
newYBASTXX0.BASTXXUHMS = time_Hms
newYBASTXX0.BASTXXUSEQ = 0


fgSelect.Visible = False
cnSab_Update.Open paramODBC_DSN_SAB

For K = 1 To fgSelect.Rows - 1
    fgSelect.Row = K
    fgSelect.Col = fgSelect_arrIndex
    meYBASTXX0 = mearrYBASTXX0(Val(fgSelect.Text))
    If meYBASTXX0.BASTXXNUM > 0 Then ' $JPL 2009-06-11$ And meYBASTXX0.BASTXXVAL <> 0 Then
        Call lstErr_ChangeLastItem(Me.lstErr, Me.cmdContext, meYBASTXX0.BASTXXDEV & " " & meYBASTXX0.BASTXXTAU)
        cmdReuters_Sab
'________________________________________________________________________________

    End If
Next K
'________________________________________________________________________________

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
    
Exit_sub:
cnSab_Update.Close
fgSelect.Visible = True
Me.Enabled = True: Me.MousePointer = 0
Call lstErr_AddItem(lstErr, cmdContext, "cmdReuters_Ok :  fin " & Time)

End Sub

Private Sub cmdSelect_Click()
Dim mFileName As String
Dim seq As Long, lenX As Long
Dim K As Integer, I As Integer
Dim X As String, XX As String
Dim xIn As String
Dim blnOk As Boolean, blnDevises As Boolean, blnEONIA As Boolean
Dim mAAMMJJ As String * 8, mHHMMSS As String * 6
Dim mAAMMJJ_VeilleO As String * 8
Dim xAAMMJJ As String * 8, xHHMMSS As String * 6
Dim mNature As String, xCodeDevise As String, xCodeTaux As String, xVal As Double
Dim wNature As String
Dim kJMA As Integer, kJJ As Integer, kMM As Integer, xMM As String, kAAAA As Integer
Dim blnVal As Boolean
On Error Resume Next


Me.Enabled = False: Me.MousePointer = vbHourglass

blnDevises = False: blnOk = False: blnEONIA = False
cmdReuters_Ok.Visible = False
fgSelect_Display
fgSelect.Visible = False
mFileName = Trim(txtSelect)
'$2007-07-23 JPL mFileName = "\\printsrvpro\BIA-REUTERS\EXPORT\Reuters-Sab.txt"
Call lstErr_Clear(lstErr, cmdContext, filDoc.FileName & ": " & Time): DoEvents

If Dir(mFileName) = "" Then
    blnError = True
    Call lstErr_AddItem(lstErr, cmdContext, filDoc.FileName & ": n'existe pas")
    GoTo Exit_sub
End If
Call lstErr_AddItem(lstErr, cmdContext, filDoc.FileName & ": " & Time): DoEvents

ReDim mearrYBASTXX0(1000)
mearrYBASTXX0_NB = 0
recYBASTXX0_Init mearrYBASTXX0(0)

Open mFileName For Input As #1
Line Input #1, xIn
K = 1
mAAMMJJ = DateJMA_Scan(xIn, K)
mAAMMJJ_VeilleO = dateElp("Ouvré", -1, mAAMMJJ)

K = InStr(K, xIn, "-") + 1
If K > 0 Then
    mHHMMSS = TimeHMS_Scan(xIn, K)
Else
    mHHMMSS = "000000"
End If
lblSelect.Caption = "Fichier export BLOOMBERG du :  " & filDoc.FileName
mNature = ""

If mAAMMJJ <> DSys And mAAMMJJ <> DSys_VeilleO Then
    blnError = Not SAB_TAU_Aut.Xspécial
    If SAB_TAU_Aut.Xspécial Then Shell_MsgBox "# Reuters import # Fichier au " & mAAMMJJ & " <> date jour" & DSys & " et " & DSys_VeilleO, vbCritical, Me.Caption, blnAuto
End If

Do Until EOF(1)
    Line Input #1, xIn
    xIn = Trim(xIn)
    If xIn <> " " Then
        seq = seq + 1
        DoEvents
        K = 0
        If Mid$(xIn, 1, 1) = ";" Then
            Call lstErr_ChangeLastItem(Me.lstErr, Me.cmdContext, xIn)
            blnDevises = False: blnOk = False: blnEONIA = False
            mNature = ""

            Select Case xIn
                Case "; COURS DEVISES": blnDevises = True
                Case "; EONIA": blnEONIA = True: mNature = "Taux"
                Case "; EOF": blnOk = True
                Case "; EURIBOR":
                    mNature = "Taux"
                Case Else
                        If Mid$(xIn, 1, 7) = "; LIBOR" Then mNature = "Taux"
                        If Mid$(xIn, 1, 6) = "; EBOR" Then mNature = "Taux"
            End Select
        Else
            xIn = xIn & " "
           If Not blnDevises And Not blnEONIA Then
                xCodeDevise = Space_Scan(xIn, K)
                xCodeTaux = Space_Scan(xIn, K)
            Else
                X = Space_Scan(xIn, K)
                If blnEONIA Then
                    xCodeDevise = "EUR"
                    If X = "EONIALAST" Then xCodeTaux = "EONIA0"
                End If
                
                If blnDevises Then
                    xCodeDevise = Mid$(X, 4, 3)
                    xCodeTaux = ""
                    Select Case Mid$(X, 7, 3)
                        Case "FIX": mNature = "Fix EUR"
                        Case "BID", "BIA":  mNature = "TC  EUR"
                        Case Else: mNature = ""
                    End Select
                End If
            End If
            X = Replace(Space_Scan(xIn, K), ",", ".")
            XX = Replace(X, ".", ",")
            If IsNumeric(XX) Then
                blnVal = True
                xVal = Val(X)
            Else
                xVal = 0
                blnVal = False
            End If
            
            xAAMMJJ = Space$(8)
            If K < Len(xIn) Then
                kJMA = InStr(K, xIn, "/")
                If kJMA > 0 Then
                    xAAMMJJ = DateJMA_Scan(xIn, K)
                Else
                    kJJ = 0: kMM = 0: kAAAA = 0: xMM = ""
                    kJMA = InStr(K, xIn, "-")
                    If kJMA > 0 Then
                        kJJ = Val(Mid$(xIn, K, kJMA - K))
                        K = kJMA + 1
                        kJMA = InStr(K, xIn, "-")
                        If kJMA > 0 Then
                            xMM = UCase$(Mid$(xIn, K, kJMA - K))
                            K = kJMA + 1
                        End If
                    Else
                        kJJ = Val(Space_Scan(xIn, K))
                        xMM = Space_Scan(xIn, K)
                        kAAAA = Val(Space_Scan(xIn, K))

                    End If
                    If kAAAA < 1000 Then kAAAA = kAAAA + 2000
                    Select Case xMM
                        Case "JAN", "JANV": kMM = 1
                        Case "FEB", "FEVR": kMM = 2
                        Case "MAR", "MARS": kMM = 3
                        Case "APR", "AVR": kMM = 4
                        Case "MAI", "MAI": kMM = 5
                        Case "JUN", "JUIN": kMM = 6
                        Case "JUL", "JUIL": kMM = 7
                        Case "AUG", "AOUT": kMM = 8
                        Case "SEP", "SEPT": kMM = 9
                        Case "OCT", "OCT": kMM = 10
                        Case "NOV", "NOV": kMM = 11
                        Case "DEC", "DEC": kMM = 12
                        Case Else: xVal = 0: kMM = 0
                    End Select
                    xAAMMJJ = Format$(kAAAA, "0000") & Format$(kMM, "00") & Format$(kJJ, "00")
                End If
                
            End If
            
            If xAAMMJJ = Space$(8) Then xAAMMJJ = mAAMMJJ
            wNature = mNature
            '$JPL 20121115
            'If xAAMMJJ <> mAAMMJJ And xAAMMJJ <> mAAMMJJ_VeilleO Then
            '    xAAMMJJ = Mid$(xAAMMJJ, 1, 4) & Mid$(xAAMMJJ, 7, 2) & Mid$(xAAMMJJ, 5, 2)
            '    If xAAMMJJ <> mAAMMJJ And xAAMMJJ <> mAAMMJJ_VeilleO Then wNature = "?date"
                
            'End If
            
            xHHMMSS = TimeHMS_Scan(xIn, K)
            
            mearrYBASTXX0_NB = mearrYBASTXX0_NB + 1
            xYBASTXX0 = mearrYBASTXX0(0)
            xYBASTXX0.BASTXXAMJ = xAAMMJJ - 19000000
            
            xYBASTXX0.BASTXXDEV = xCodeDevise
            xYBASTXX0.BASTXXTAU = xCodeTaux
            xYBASTXX0.BASTXXVAL = xVal
            
            If Not blnVal Then
                xYBASTXX0.BASTXXNUM = 0
                wNature = "?valeur"
            Else
                 Select Case wNature
                     Case "Taux"
                                     xYBASTXX0.BASTXXNUM = 25
                                     'xYBASTXX0.BASTXXARG = xCodeDevise & xCodeTaux
                     Case "Fix EUR"
                                     xYBASTXX0.BASTXXNUM = 37
                                     'xYBASTXX0.BASTXXARG = xCodeDevise & xCodeTaux
                                     'xYBASTXX0.BASTXXDEV2 = "EUR"
                      Case Else
                                     xYBASTXX0.BASTXXNUM = 0
                End Select
            End If
            
            mearrYBASTXX0(mearrYBASTXX0_NB) = xYBASTXX0

            
            fgSelect.Rows = fgSelect.Rows + 1
            fgSelect.Row = fgSelect.Rows - 1
          
            fgSelect.Col = 0: fgSelect.Text = wNature
            fgSelect.Col = 1: fgSelect.Text = xCodeDevise
            
            fgSelect.Col = 2: fgSelect.Text = xCodeTaux
            
            fgSelect.Col = 3: fgSelect.Text = xVal
            fgSelect.Col = 4: fgSelect.Text = dateImp10(xAAMMJJ) & "  " & timeImp8(xHHMMSS)
            fgSelect.Col = fgSelect_arrIndex: fgSelect.Text = mearrYBASTXX0_NB
         
         End If
    End If
        
Loop
Close
ReDim Preserve mearrYBASTXX0(mearrYBASTXX0_NB + 1)
fgSelect.Visible = True
Call lstErr_AddItem(frmSAB_TAU.lstErr, frmSAB_TAU.cmdContext, filDoc.FileName & " , nb : " & mearrYBASTXX0_NB)
If Not blnOk Then
    blnError = True
    Shell_MsgBox "frmSAB_TAU.cmdSelect#  manque 'EOF'", vbCritical, Me.Caption, blnAuto
 
End If
If Not blnError Then cmdReuters_Ok.Visible = blncmdReuters_Ok_Visible

Exit_sub:

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
blncmdReuters_Ok_Visible = SAB_TAU_Aut.Xspécial
txtSelect = paramSAB_TAU_Archive & filDoc.FileName
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

If fraContextOptions.Visible Then fraContextOptions.Visible = False: Exit Sub

If SSTab1.Tab = 0 Then
    If cmdReuters_Ok.Visible Then
        cmdReuters_Ok.Visible = False
        fgSelect.Clear
    Else
        Unload Me
    End If
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
lblContextOptions.Caption = "Sélectionner un fichier" & Asc10_13 & "touche (Esc) pour abandonner"
lblContextOptions.ForeColor = warnUsrColor
filDoc.PATH = paramSAB_TAU_Archive
filDoc.Pattern = "*.XXX"
Call DTPicker_Control(txtSelect_AMJ, wAMJMin)
filDoc.Pattern = wAMJMin & "*.*"
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





Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim wOrigine As String
On Error Resume Next
If y <= fgSelect.RowHeightMin Then
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
        meYBASTXX0 = mearrYBASTXX0(Val(fgSelect.Text))

        If meYBASTXX0.BASTXXNUM > 0 And meYBASTXX0.BASTXXVAL <> 0 Then Me.PopupMenu mnufgSelect, vbPopupMenuLeftButton
   End If
End If
End Sub

Public Sub fgSelect_Reset()
fgSelect.Clear
fgSelect_Sort1 = 0: fgSelect_Sort2 = 0
fgSelect_Sort1_Old = -1
fgSelect_RowDisplay = 0: fgSelect_RowClick = 0
fgSelect_arrIndex = 7
blnfgSelect_DisplayLine = False
fgSelect_SortAD = 6
fgSelect.LeftCol = 0

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





Private Sub txtSelect_AMJ_Change()
fgSelect.Clear
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

Public Function cmdReuters_Sab()
Dim V
newYBASTXX0.BASTXXUSEQ = newYBASTXX0.BASTXXUSEQ + 1
newYBASTXX0.BASTXXNUM = meYBASTXX0.BASTXXNUM
newYBASTXX0.BASTXXDEV = meYBASTXX0.BASTXXDEV
newYBASTXX0.BASTXXTAU = meYBASTXX0.BASTXXTAU
newYBASTXX0.BASTXXAMJ = meYBASTXX0.BASTXXAMJ
newYBASTXX0.BASTXXVAL = meYBASTXX0.BASTXXVAL

'If InStr(newYBASTXX0.BASTXXVAL, "-") > 0 Then
'    cmdSendMail_ALERTE_TAU_Négatif
'    newYBASTXX0.BASTXXVAL = 0
'End If
V = sqlYBASTXX0_Insert(newYBASTXX0)
fgSelect.Col = 5
If IsNull(V) Then
    fgSelect_ForeColor lblUsr.ForeColor
    If Not blnUPDATE_FIX And newYBASTXX0.BASTXXNUM = 37 And newYBASTXX0.BASTXXDEV = "USD" And newYBASTXX0.BASTXXAMJ = DSYS_IBM Then
        blnCONTROL_FIX = True
        newEURUSD = newYBASTXX0.BASTXXVAL
    End If
    If Not blnUPDATE_TAU And newYBASTXX0.BASTXXNUM = 25 And newYBASTXX0.BASTXXDEV = "EUR" And newYBASTXX0.BASTXXTAU = "EURJ1M" And newYBASTXX0.BASTXXAMJ = DSYS_IBM Then
        blnCONTROL_TAU = True
        newEURJ1M = newYBASTXX0.BASTXXVAL
    End If

Else
    fgSelect_ForeColor vbRed
    fgSelect.Text = V
    'Shell_MsgBox "mnufgSelect_AddNew#  " & meYBASTXX0.Err, vbCritical, Me.Caption, False
End If

End Function

Public Sub xxx_auto(lFile As String)
txtSelect = paramSAB_TAU_Archive & lFile
cmdSelect_Click
cmdReuters_Ok_Click

End Sub

Public Sub ZBASTAB0_Load(lAMJ As String, blnInit As Boolean)
On Error GoTo Error_Handler
Dim xSQL As String, xAMJ As String, wCours As Double
Dim impAMJ As String, xArg As String
Dim blnNormal As Boolean, wAmj As String


Dim wA1 As Integer, wA2 As Integer, wA3 As Integer, wA4 As Integer
On Error GoTo Error_Handler
fgSelect_Display
fgSelect.Visible = False


                impAMJ = "   " & dateImp10(lAMJ)
                lblSelect = "Cours SAB au " & impAMJ
                
                    
                 xAMJ = dateIBM(lAMJ)
                 Call convX2P_IBMAMJ(xAMJ, wA1, wA2, wA3, wA4)
                    
                xSQL = "select * from " & paramIBM_Library_SAB & ".ZBASTAB0" _
                     & " where BASTABETA = " & currentZMNURUT0.MNURUTETB _
                     & " and   BASTABNUM = 37" _
                     & " and ascii(substring(bastablo1 , 1 , 1)) = " & wA1 _
                     & " and ascii(substring(bastablo1 , 2 , 1)) = " & wA2 _
                     & " and ascii(substring(bastablo1 , 3 , 1)) = " & wA3 _
                     & " and ascii(substring(bastablo1 , 4 , 1)) = " & wA4 _
                     & " order by BASTABARG"

Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    X = rsSab("BASTABARG")
    wAmj = 19000000 + convX2P(Mid$(rsSab("BASTABARG"), 4, 4))
    If wAmj = lAMJ Then
        fgSelect.Rows = fgSelect.Rows + 1
        fgSelect.Row = fgSelect.Rows - 1
        fgSelect.Col = 0: fgSelect.Text = "SAB_COURS"
        xArg = Mid$(X, 1, 3)
        fgSelect.Col = 1: fgSelect.Text = xArg
        
        wCours = CDbl(convX2P(Mid$(rsSab("BASTABDON"), 1, 8))) / 1000000000
        fgSelect.Col = 3: fgSelect.Text = wCours
        fgSelect.Col = 4: fgSelect.Text = dateImp10(wAmj) 'impAMJ
        If blnInit And xArg = "USD" Then
            blnEURUSD = True
            sabEURUSD = wCours
        End If
    End If
    rsSab.MoveNext

Loop
''_________

'     & " and   BASTABARG like '_________" & xAMJ & "%'" _


xSQL = "select * from " & paramIBM_Library_SAB & ".ZBASTAB0" _
     & " where BASTABETA = " & currentZMNURUT0.MNURUTETB _
     & " and   BASTABNUM = 25" _
    & " and ascii(substring(bastabarg , 10 , 1)) = " & wA1 _
    & " and ascii(substring(bastabarg , 11 , 1)) = " & wA2 _
    & " and ascii(substring(bastabarg , 12 , 1)) = " & wA3 _
    & " and ascii(substring(bastabarg , 13 , 1)) = " & wA4 _
     & " order by BASTABARG"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    X = rsSab("BASTABARG")
    wAmj = 19000000 + convX2P(Mid$(rsSab("BASTABARG"), 10, 4))
    If wAmj = lAMJ Then

        fgSelect.Rows = fgSelect.Rows + 1
        fgSelect.Row = fgSelect.Rows - 1
        fgSelect.Col = 0: fgSelect.Text = "SAB_TAUX"
        fgSelect.Col = 1: fgSelect.Text = Mid$(X, 1, 3)
        fgSelect.Col = 2: fgSelect.Text = Mid$(X, 4, 6)
        
        wCours = CDbl(convX2P(Mid$(rsSab("BASTABDON"), 1, 8))) / 1000000000
       fgSelect.Col = 3: fgSelect.Text = wCours
       If wCours < 0 Then fgSelect.CellBackColor = vbRed
        fgSelect.Col = 4: fgSelect.Text = dateImp10(wAmj)
        If blnInit And Mid$(X, 1, 9) = "EUREURJ1M" Then
            blnEURJ1M = True
            sabEURJ1M = wCours
        End If
    End If

    rsSab.MoveNext

Loop
'==========================================================================
                lblSelect = "Taux, Cours et Cours Espèces au " & impAMJ
                 xAMJ = dateIBM(lAMJ)
                 Call convX2P_IBMAMJ(xAMJ, wA1, wA2, wA3, wA4)
                xSQL = "select * from " & paramIBM_Library_SAB & ".ZBASTAB0" _
                     & " where BASTABETA = " & currentZMNURUT0.MNURUTETB _
                     & " and   BASTABNUM = 38" _
                     & " and ascii(substring(bastabarg , 6 , 1)) = " & wA1 _
                     & " and ascii(substring(bastabarg , 7 , 1)) = " & wA2 _
                     & " and ascii(substring(bastabarg , 8 , 1)) = " & wA3 _
                     & " and ascii(substring(bastabarg , 9 , 1)) = " & wA4 _
                     & " order by BASTABARG"
Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    X = rsSab("BASTABARG")
    wAmj = 19000000 + convX2P(Mid$(rsSab("BASTABARG"), 6, 4))
    If wAmj = lAMJ Then
        fgSelect.Rows = fgSelect.Rows + 1
        fgSelect.Row = fgSelect.Rows - 1
        fgSelect.Col = 0: fgSelect.Text = "ESPECES"
        xArg = Mid$(X, 3, 3)
        fgSelect.Col = 1: fgSelect.Text = xArg
        
        fgSelect.Col = 2: fgSelect.Text = Mid$(X, 33, 3)
        
        wCours = CDbl(convX2P(Mid$(rsSab("BASTABDON"), 1, 8))) / 1000000000
        fgSelect.Col = 3: fgSelect.Text = wCours
        fgSelect.Col = 4: fgSelect.Text = dateImp10(wAmj)
    End If
    rsSab.MoveNext
Loop
'==========================================================================

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
    
Exit_sub:
fgSelect.Visible = True
Me.Enabled = True: Me.MousePointer = 0
Call lstErr_AddItem(lstErr, cmdContext, "cmdReuters_Ok :  fin " & Time)

End Sub

Public Sub paramInit_Auto()

blnCONTROL_TAU = False
blnCONTROL_FIX = False

blnALERTE_TAU = False
xYBIAMON0.MONAPP = "@SAB_TAU"
xYBIAMON0.MONFLUX = "ALERTE_TAU"
V = rsYBIAMON0_Read(xYBIAMON0)
If IsNull(V) And Trim(xYBIAMON0.MONFILE) = DSys Then blnALERTE_TAU = True

blnALERTE_FIX = False
xYBIAMON0.MONAPP = "@SAB_TAU"
xYBIAMON0.MONFLUX = "ALERTE_FIX"
V = rsYBIAMON0_Read(xYBIAMON0)
If IsNull(V) And Trim(xYBIAMON0.MONFILE) = DSys Then blnALERTE_FIX = True


blnUPDATE_TAU = False
xYBIAMON0.MONAPP = "@SAB_TAU"
xYBIAMON0.MONFLUX = "UPDATE_TAU"
V = rsYBIAMON0_Read(xYBIAMON0)
If IsNull(V) And Trim(xYBIAMON0.MONFILE) = DSys Then blnUPDATE_TAU = True: blnALERTE_TAU = True

blnUPDATE_FIX = False
xYBIAMON0.MONAPP = "@SAB_TAU"
xYBIAMON0.MONFLUX = "UPDATE_FIX"
V = rsYBIAMON0_Read(xYBIAMON0)
If IsNull(V) And Trim(xYBIAMON0.MONFILE) = DSys Then blnUPDATE_FIX = True: blnALERTE_FIX = True

End Sub

Public Sub ZBASTAB0_Control()
Dim V
Dim blnUPDATE_Mail As Boolean
Dim wSendMail As typeSendMail
Dim xDétail As String
On Error Resume Next


blnUPDATE_Mail = False

Call ZBASTAB0_Load(DSys, True)

If blnCONTROL_FIX Then
    wSendMail.Subject = dateImp10(DSys) & "cours EUR /USD : " & newEURUSD
    xYBIAMON0.MONAPP = "@SAB_TAU"
    xYBIAMON0.MONFLUX = "UPDATE_FIX"
    V = rsYBIAMON0_Read(xYBIAMON0)
    If IsNull(V) And sabEURUSD = newEURUSD Then
        newYBIAMON0 = xYBIAMON0
        newYBIAMON0.MONFILE = DSys
        cnSab_Update.Open paramODBC_DSN_SAB
        Call fctExploitation_Transaction_End(newYBIAMON0, xYBIAMON0)
        cnSab_Update.Close
        blnCONTROL_FIX = False
        blnUPDATE_FIX = True
        blnALERTE_FIX = True
        blnUPDATE_Mail = True
    End If
End If
If blnCONTROL_TAU Then
    wSendMail.Subject = dateImp10(DSys) & "Taux EURJ1M : " & newEURJ1M
    xYBIAMON0.MONAPP = "@SAB_TAU"
    xYBIAMON0.MONFLUX = "UPDATE_TAU"
    V = rsYBIAMON0_Read(xYBIAMON0)
    If IsNull(V) And sabEURJ1M = newEURJ1M Then
        newYBIAMON0 = xYBIAMON0
        newYBIAMON0.MONFILE = DSys
        cnSab_Update.Open paramODBC_DSN_SAB
        Call fctExploitation_Transaction_End(newYBIAMON0, xYBIAMON0)
        cnSab_Update.Close
        blnCONTROL_TAU = False
        blnUPDATE_TAU = True
        blnALERTE_TAU = True
        blnUPDATE_Mail = True
    End If
End If

If blnUPDATE_Mail Then
    Call cmdSendMail_Détail(xDétail)
    
    wSendMail.FromDisplayName = "UPDATE_TAU"
    wSendMail.RecipientDisplayName = "@SAB_TAU"
    
    wSendMail.Attachment = ""
    
    '#87CEFA
    wSendMail.Message = "<body bgcolor = #FFFFFF>" _
                        & "<BR>" & xDétail
    
    wSendMail.AsHTML = True
    
    srvSendMail.Monitor wSendMail
End If
End Sub

Public Sub cmdSendMail_ALERTE_TAU()
Dim wSendMail As typeSendMail
Dim xDétail As String

On Error Resume Next


xYBIAMON0.MONAPP = "@SAB_TAU"
xYBIAMON0.MONFLUX = "ALERTE_TAU"
V = rsYBIAMON0_Read(xYBIAMON0)
If IsNull(V) Then
    newYBIAMON0 = xYBIAMON0
    newYBIAMON0.MONFILE = DSys
    cnSab_Update.Open paramODBC_DSN_SAB
    Call fctExploitation_Transaction_End(newYBIAMON0, xYBIAMON0)
    cnSab_Update.Close
    blnALERTE_TAU = True
End If

Call cmdSendMail_Détail(xDétail)

wSendMail.FromDisplayName = "ALERTE_TAU"
wSendMail.RecipientDisplayName = "@SAB_TAU"

wSendMail.Subject = dateImp10(DSys) & " " & Time & " , le taux EURJ1M n'est pas mis à jour dans SAB, en provenance de BLOOMBERG"
wSendMail.Attachment = ""
wSendMail.Message = "<body bgcolor=" & Asc34 & "MAGENTA" & Asc34 & ">" _
                    & "<BR>" & xDétail

wSendMail.AsHTML = True

srvSendMail.Monitor wSendMail

End Sub
Public Sub cmdSendMail_ALERTE_TAU_Négatif()
Dim wSendMail As typeSendMail
Dim xDétail As String

On Error Resume Next


xYBIAMON0.MONAPP = "@SAB_TAU"
xYBIAMON0.MONFLUX = "ALERTE_TAU"

wSendMail.FromDisplayName = "ALERTE_TAU"
wSendMail.RecipientDisplayName = "@SAB_TAU"
xDétail = "Le taux " _
                  & newYBASTXX0.BASTXXDEV & "  " _
                  & newYBASTXX0.BASTXXTAU & "  " _
                  & "en provenance de BLOOMBERG est négatif : " _
                  & newYBASTXX0.BASTXXVAL & "  " _
                  & "en date du " & dateIBM10(newYBASTXX0.BASTXXAMJ, True)
wSendMail.Subject = xDétail
wSendMail.Attachment = ""
wSendMail.Message = "<body bgcolor=" & Asc34 & "RED" & Asc34 & ">" _
                    & "<BR>" & xDétail & "<BR>" & "le taux est initialisé à ZERO dans SAB"
wSendMail.AsHTML = True

srvSendMail.Monitor wSendMail

End Sub

Public Sub cmdSendMail_ALERTE_FIX()
Dim wSendMail As typeSendMail
Dim xDétail As String
On Error Resume Next

    xYBIAMON0.MONAPP = "@SAB_TAU"
    xYBIAMON0.MONFLUX = "ALERTE_FIX"
    V = rsYBIAMON0_Read(xYBIAMON0)
    If IsNull(V) Then
        newYBIAMON0 = xYBIAMON0
        newYBIAMON0.MONFILE = DSys
        cnSab_Update.Open paramODBC_DSN_SAB
        Call fctExploitation_Transaction_End(newYBIAMON0, xYBIAMON0)
        cnSab_Update.Close
        blnALERTE_FIX = True
    End If
    Call cmdSendMail_Détail(xDétail)
    wSendMail.FromDisplayName = "ALERTE_FIX"
    wSendMail.RecipientDisplayName = "@SAB_TAU"
    wSendMail.Subject = dateImp10(DSys) & " " & Time & " , le cours EUR / USD n'est pas mis à jour dans SAB, en provenance de BLOOMBERG"
    wSendMail.Attachment = ""
    wSendMail.Message = "<body bgcolor=" & Asc34 & "MAGENTA" & Asc34 & ">" _
                        & "<BR>" & xDétail
    wSendMail.AsHTML = True
    srvSendMail.Monitor wSendMail

End Sub


Public Sub cmdSendMail_Détail(lDétail As String)
Dim xDevise As String, xCodeTaux As String, xValeur As String, xDate As String
Dim K As Integer


Call ZBASTAB0_Load(DSys, True)

lDétail = "<TABLE border = 1  width=1000 height=5 bgcolor=#0000FF cellpadding=4 ><TR>" _
         & "<TD  width=200 height=5><span style='font-size:10.0pt;font-family:Arial'><Font color=#FFFFFF>Devise</TD>" _
         & "<TD  width=200 height=5><span style='font-size:10.0pt;font-family:Arial'><Font color=#FFFFFF>Taux</B></TD>" _
         & "<TD  width=400 height=5><span style='font-size:10.0pt;font-family:Arial'><Font color=#FFFFFF>Valeur</TD>" _
         & "<TD  width=200 height=5><span style='font-size:10.0pt;font-family:Arial'><Font color=#FFFFFF>Date</TD>" _
        & "</TR></TABLE>"



For K = 1 To fgSelect.Rows - 1
    fgSelect.Row = K
    fgSelect.Col = 1: xDevise = fgSelect.Text
    fgSelect.Col = 2: xCodeTaux = fgSelect.Text
    fgSelect.Col = 3: xValeur = fgSelect.Text
    fgSelect.Col = 4: xDate = fgSelect.Text

    lDétail = lDétail & "<TABLE   width=1000 border=1 cellpadding=4 ></B><TR>" _
         & "<TD " & "bgcolor = #87FAFA" & " width=200 height=5><span style='font-size:8.0pt;font-family:Arial'>" & htmlFontColor_Blue & xDevise & "</TD>" _
         & "<TD " & "bgcolor = #87FAFA" & " width=200 height=5><span style='font-size:8.0pt;font-family:Arial'>" & htmlFontColor_Blue & xCodeTaux & "</TD>" _
         & "<TD " & "bgcolor = #87FAFA" & " width=400 height=5><span style='font-size:8.0pt;font-family:Arial'>" & htmlFontColor_Blue & xValeur & "</TD>" _
         & "<TD " & "bgcolor = #87FAFA" & " width=200 height=5><span style='font-size:8.0pt;font-family:Arial'>" & xDate & "</TD>" _
         & "</TD></TR></TABLE>"

Next K
End Sub
