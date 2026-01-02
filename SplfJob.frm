VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmSPLFJOB 
   AutoRedraw      =   -1  'True
   Caption         =   "SPLFJOB : Gestion des spoules"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13875
   Icon            =   "SplfJob.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9495
   ScaleWidth      =   13875
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   450
      Left            =   7200
      TabIndex        =   7
      Top             =   0
      Width           =   6135
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8895
      Left            =   45
      TabIndex        =   2
      Top             =   525
      Width           =   13770
      _ExtentX        =   24289
      _ExtentY        =   15690
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "NT"
      TabPicture(0)   =   "SplfJob.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraEditionNT"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "AS400"
      TabPicture(1)   =   "SplfJob.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraEditionAS400"
      Tab(1).Control(1)=   "fraEditionAS400_Import"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Informatique"
      TabPicture(2)   =   "SplfJob.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdZXSWITG2"
      Tab(2).Control(1)=   "cmdJPL"
      Tab(2).Control(2)=   "fgSelect"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "FTP"
      TabPicture(3)   =   "SplfJob.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraFTP"
      Tab(3).ControlCount=   1
      Begin VB.CommandButton cmdZXSWITG2 
         BackColor       =   &H00FF0000&
         Caption         =   "Conversion T2DIR (""C:\Temp\sab\ZXSWITG2.csv"" = > ""C:\Temp\sab\ZXSWITG2.txt"" ) "
         Height          =   1332
         Left            =   -68760
         TabIndex        =   35
         Top             =   840
         Width           =   3375
      End
      Begin VB.CommandButton cmdJPL 
         BackColor       =   &H00FF0000&
         Caption         =   "JPL TEST"
         Height          =   1575
         Left            =   -74280
         TabIndex        =   34
         Top             =   1320
         Width           =   3375
      End
      Begin VB.Frame fraFTP 
         Height          =   8430
         Left            =   -74910
         TabIndex        =   18
         Top             =   405
         Width           =   13515
         Begin VB.CommandButton cmdFTP_NT_AS400 
            BackColor       =   &H0080C0FF&
            Caption         =   "FTP : NT => AS400 "
            Height          =   1695
            Left            =   10200
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   360
            Width           =   2775
         End
         Begin VB.CommandButton cmdFTP_AS400_NT 
            BackColor       =   &H0080FF80&
            Caption         =   "FTP : AS400 => NT"
            Height          =   1710
            Left            =   210
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   270
            Width           =   3135
         End
         Begin VB.Frame fraFTP_Nt 
            Caption         =   "NT"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1140
            Left            =   3960
            TabIndex        =   25
            Top             =   5400
            Width           =   9120
            Begin VB.TextBox txtFTP_NT_File 
               Height          =   285
               Left            =   1320
               TabIndex        =   27
               Text            =   "D:\Temp\FTP\"
               Top             =   480
               Width           =   7200
            End
            Begin VB.Label lblFTP_NT_File 
               Caption         =   "Fichier "
               Height          =   270
               Left            =   240
               TabIndex        =   26
               Top             =   480
               Width           =   600
            End
         End
         Begin VB.Frame fraFTP_AS400 
            Caption         =   "AS400"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4005
            Left            =   4440
            TabIndex        =   19
            Top             =   480
            Width           =   4305
            Begin VB.ComboBox cboFTP_AS400_File 
               Height          =   315
               Left            =   1920
               Sorted          =   -1  'True
               TabIndex        =   31
               Text            =   "cboFTP_AS400_File"
               Top             =   2160
               Width           =   1755
            End
            Begin VB.ComboBox cboFTP_AS400_Library 
               Height          =   315
               Left            =   1920
               Sorted          =   -1  'True
               TabIndex        =   30
               Text            =   "cboFTP_AS400_Library"
               Top             =   1320
               Width           =   1755
            End
            Begin VB.ComboBox cboFTP_AS400_host 
               Height          =   315
               Left            =   1920
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   24
               Top             =   480
               Width           =   1755
            End
            Begin VB.CheckBox chkFTP_AS400_Binary 
               Alignment       =   1  'Right Justify
               Caption         =   "Binary"
               Height          =   315
               Left            =   600
               TabIndex        =   23
               Top             =   3000
               Width           =   1155
            End
            Begin VB.Label lblFTP_AS400_File 
               Caption         =   "fichier"
               Height          =   315
               Left            =   600
               TabIndex        =   22
               Top             =   2280
               Width           =   555
            End
            Begin VB.Label lblFTP_AS400_Library 
               Caption         =   "Librairie"
               Height          =   315
               Left            =   600
               TabIndex        =   21
               Top             =   1320
               Width           =   690
            End
            Begin VB.Label lblFTP_AS400_host 
               Caption         =   "Host"
               Height          =   315
               Left            =   600
               TabIndex        =   20
               Top             =   480
               Width           =   660
            End
         End
      End
      Begin VB.Frame fraEditionAS400_Import 
         Height          =   1365
         Left            =   -74880
         TabIndex        =   11
         Top             =   7320
         Width           =   13620
         Begin VB.CommandButton cmdSplfMonitor_Ftp 
            Caption         =   "Lecture SPLFFTPW0"
            Height          =   735
            Left            =   10680
            TabIndex        =   14
            Top             =   360
            Width           =   2130
         End
         Begin VB.TextBox txtSplfMonitor_Import 
            Height          =   285
            Left            =   2625
            TabIndex        =   13
            Text            =   "D:\Temp\ .txt"
            Top             =   345
            Width           =   6585
         End
         Begin VB.TextBox txtSplfMonitor_Export 
            Height          =   285
            Left            =   2625
            TabIndex        =   12
            Text            =   "D:\Temp\ .txt"
            Top             =   825
            Width           =   6600
         End
         Begin VB.Label lblSplfMonitor_Import 
            Caption         =   "fichier d'import SPLFFTPW0"
            Height          =   255
            Left            =   135
            TabIndex        =   16
            Top             =   375
            Width           =   2025
         End
         Begin VB.Label lblSplfMonitor_Export 
            Caption         =   "Répertoire de destination"
            Height          =   255
            Left            =   75
            TabIndex        =   15
            Top             =   885
            Width           =   1920
         End
      End
      Begin VB.Frame fraEditionAS400 
         Height          =   6615
         Left            =   -74880
         TabIndex        =   10
         Top             =   480
         Width           =   13455
         Begin VB.CommandButton cmdSplfMonitor_Dir 
            BackColor       =   &H00C0FFC0&
            Caption         =   "S820i_Out\SPLF\ => NT"
            Height          =   495
            Left            =   7200
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   1200
            Width           =   2175
         End
         Begin VB.FileListBox filDoc 
            Height          =   5745
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   6735
         End
      End
      Begin VB.Frame fraEditionNT 
         Height          =   8415
         Left            =   105
         TabIndex        =   3
         Top             =   360
         Width           =   13515
         Begin VB.CommandButton cmdSplf_NoPaper_Clear 
            BackColor       =   &H00FF00FF&
            Caption         =   "Effacer les anciens répertoires  NoPaper "
            Height          =   840
            Left            =   9765
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   7080
            Width           =   3135
         End
         Begin VB.CommandButton cmdSplf_Corbeille 
            BackColor       =   &H00C0C0FF&
            Caption         =   "Mettre à la corbeille les documents dont la date de création est  < **/**/****"
            Height          =   1290
            Left            =   9675
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   3915
            Width           =   3120
         End
         Begin VB.CommandButton cmdSplf_Corbeille_Clear 
            BackColor       =   &H000000FF&
            Caption         =   "Effacer de la corbeille les documents dont la date de création est  < **/**/****"
            Height          =   1395
            Left            =   9720
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   5325
            Width           =   3135
         End
         Begin VB.DirListBox dirW 
            Height          =   2790
            Left            =   9120
            TabIndex        =   6
            Top             =   240
            Width           =   4260
         End
         Begin VB.FileListBox filW 
            Height          =   7500
            Left            =   135
            TabIndex        =   5
            Top             =   225
            Width           =   8550
         End
         Begin MSComCtl2.DTPicker txtSplfAmj 
            Height          =   300
            Left            =   10530
            TabIndex        =   4
            Top             =   3390
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   393216
            CalendarBackColor=   16777215
            CalendarForeColor=   0
            CalendarTitleBackColor=   8421504
            CalendarTitleForeColor=   16777215
            CalendarTrailingForeColor=   12632256
            CustomFormat    =   "dd  MM yyy"
            Format          =   125698051
            CurrentDate     =   36299
            MaxDate         =   401768
            MinDate         =   -328351
         End
      End
      Begin MSFlexGridLib.MSFlexGrid fgSelect 
         Height          =   1140
         Left            =   -74730
         TabIndex        =   17
         Top             =   4200
         Width           =   8730
         _ExtentX        =   15399
         _ExtentY        =   2011
         _Version        =   393216
         Rows            =   1
         Cols            =   7
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
         FormatString    =   $"SplfJob.frx":037A
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
      Left            =   13320
      Picture         =   "SplfJob.frx":047D
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   500
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
End
Attribute VB_Name = "frmSPLFJOB"
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
Dim SplfJobAut As typeAuthorization
Dim blnAuto_SplfJob_Run As Boolean

Dim fgSelect_FormatString As String, fgSelect_K As Integer
Dim fgSelect_RowDisplay As Integer, fgSelect_RowClick As Integer, fgSelect_ColClick As Integer
Dim fgSelect_ColorClick As Long, fgSelect_ColorDisplay As Long
Dim fgSelect_Sort1 As Integer, fgSelect_Sort2 As Integer
Dim fgSelect_SortAD As Integer, fgSelect_Sort1_Old As Integer
Dim fgSelect_arrIndex As Integer
Dim blnfgSelect_DisplayLine As Boolean

Dim meSplfJob As typeSplfJob
Dim blnError As Boolean
Dim blnAuto_SplfJob As Boolean
Dim wSplfAmj As String * 8
Dim Nb As Long, Nb_Ftp As Long
Dim fsoFile As File
Dim blnFTP_Get As Boolean
Dim paramNT_Folder As String


Dim xUser As typeUser, xEdition_Form As typeEdition_Form

Public cnAdo As New ADODB.Connection
Public rsado As New ADODB.Recordset

Private Sub fgSelect_Display()

SSTab1.Tab = 1

fgSelect.Visible = True
fgSelect.Clear: fgSelect.Rows = 1: fgSelect_RowDisplay = 0

fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
fgSelect.Enabled = True



fgSelect_SortAD = 6
If fgSelect.Rows > 1 Then fgSelect_Sort
fgSelect.LeftCol = 0

End Sub

Public Sub fgSelect_DisplayLine()
On Error Resume Next

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
    Select Case lK
        Case 1, 2: fgSelect.Text = X
    End Select
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

Call BiaPgmAut_Init(Mid$(Msg, 1, 12), SplfJobAut)

'Call charge_imprimantes

'blnSetfocus = True
Form_Init

Select Case UCase$(Trim(Mid$(Msg, 1, 12)))
    Case "@AUTO_SPLF":     blnAuto_SplfJob = True: Auto_SplfJob
    Case "@AUTO_CLROUT":   blnAuto_SplfJob = True: mainSoc_AMJCPT_Load: Auto_CLROUTQ
   Case Else: blnAuto_SplfJob = False
End Select

End Sub

Sub charge_imprimantes()
Dim ii As Long
Dim z As String
Dim fic As Long
Dim ligIn As String
Dim collimp As String

ReDim collection_IMP(0)

Call rsElpTable_Read("Param", "collection", "IMPRIMANTES", "", collimp)
If collimp <> "" Then
    fic = FreeFile
    If Right(paramFolder_Local, 1) <> "\" Then
        z = paramFolder_Local & "\"
    Else
        z = paramFolder_Local
    End If
    ii = 0
    Open z & collimp For Input As #fic
    Do Until EOF(fic)
        Line Input #fic, ligIn
        ii = ii + 1
    Loop
    Close #fic
    'Le nombre total d'imprimantes se trouve dans le poste 1
    ReDim collection_IMP(ii + 1)
    collection_IMP(1) = CStr(ii)
    ii = 1
    Open z & collimp For Input As #fic
    Do Until EOF(fic)
        Line Input #fic, ligIn
        ii = ii + 1
        collection_IMP(ii) = Trim(ligIn)
    Loop
    Close #fic
End If
End Sub

Public Sub Form_Init()
Me.Enabled = False
Call lstErr_Clear(lstErr, cmdPrint, "Initialisation ")
DoEvents
dirW.Enabled = False
filW.Enabled = False

If Not IsNull(param_Init) Then
    MsgBox "paramétrage inconsistant", vbCritical, "frmSPLFJOB.param_init"
    Unload Me
End If

blnControl = False
fgSelect_FormatString = fgSelect.FormatString
fgSelect.Enabled = True

cmdReset
Me.Enabled = True

End Sub


'---------------------------------------------------------
Public Sub cmdReset()
'---------------------------------------------------------
Dim X As String
blnControl = False
usrColor_Set
cmdContext.Caption = constcmdRechercher: blnMsgBox_Quit = False
arrTag_Set False
lstErr.Visible = False
currentAction = ""

cmdSplf_Corbeille.Enabled = SplfJobAut.Xspécial
cmdSplf_Corbeille_Clear.Enabled = SplfJobAut.Xspécial

filDoc.Visible = False
txtSplfMonitor_Import = paramEditionFtp_File
txtSplfMonitor_Export = paramEditionSplf_Folder


Call DTPicker_Set(txtSplfAmj, DSys_VeilleO)
blncmdSplfMonitor = False

cboFTP_AS400_host.Clear
cboFTP_AS400_host.AddItem "P6A8"
cboFTP_AS400_host.AddItem "SABA"
cboFTP_AS400_host.ListIndex = 1

cboFTP_AS400_Library.Clear
cboFTP_AS400_Library.AddItem ""
cboFTP_AS400_Library.AddItem paramIBM_Library_SAB
cboFTP_AS400_Library.AddItem paramIBM_Library_SABSPE
cboFTP_AS400_Library.AddItem paramIBM_Library_File
cboFTP_AS400_Library.AddItem paramIBM_Library_Src
cboFTP_AS400_Library.AddItem paramIBM_Library_Obj
cboFTP_AS400_Library.AddItem "SAB073U"
cboFTP_AS400_Library.AddItem "SAB073USPE"
cboFTP_AS400_Library.ListIndex = 2

paramNT_Folder = "C:\Temp\"

cboFTP_AS400_File.Clear
cboFTP_AS400_File.AddItem ""
cboFTP_AS400_File.AddItem "DSPFFDY0"
cboFTP_AS400_File.AddItem "DSPFDY0"
cboFTP_AS400_File.AddItem "DSPFDY1"
cboFTP_AS400_File.AddItem "DSPFDY2"
cboFTP_AS400_File.AddItem "ELPKMPGM"
cboFTP_AS400_File.AddItem "YMNUMEN0"
cboFTP_AS400_File.AddItem "YMNURUT0"
cboFTP_AS400_File.AddItem "YMNUOPT0"
cboFTP_AS400_File.AddItem "YMNUUTI0"
cboFTP_AS400_File.AddItem "YPLAN0"
cboFTP_AS400_File.AddItem "YCLIENA0"
cboFTP_AS400_File.AddItem "YBASTAU0"
cboFTP_AS400_File.AddItem "YCHEOPP0"
cboFTP_AS400_File.AddItem "YCHQCOM0"
cboFTP_AS400_File.AddItem "YCHQHIS0"
cboFTP_AS400_File.AddItem "YCOMPTE0"
cboFTP_AS400_File.AddItem "YSITLOT0"
cboFTP_AS400_File.AddItem "YSITNUM0"
cboFTP_AS400_File.AddItem "YSITORD0"
cboFTP_AS400_File.AddItem "YSITPR10"
cboFTP_AS400_File.AddItem "YSITPR20"
cboFTP_AS400_File.AddItem "YMNUETA0"
cboFTP_AS400_File.AddItem "YMOUVEA0"
cboFTP_AS400_File.AddItem "YLIBEL0"
cboFTP_AS400_File.AddItem "YSOLDE0"
cboFTP_AS400_File.AddItem "YBIAMVT0"
cboFTP_AS400_File.AddItem "YTITULA0"
cboFTP_AS400_File.AddItem "YREPCPT0"
cboFTP_AS400_File.AddItem "YAUTE1I0"

cboFTP_AS400_File.AddItem "YCGSENC0"
cboFTP_AS400_File.AddItem "YCGSCOM0"
cboFTP_AS400_File.AddItem "YCGSMOY0"
cboFTP_AS400_File.AddItem "YCGSMM10"
cboFTP_AS400_File.AddItem "YCGSMM30"

'cboFTP_AS400_File.AddItem "YXXXXXX0"

cboFTP_AS400_File.ListIndex = 0


blnControl = True

End Sub


Public Function param_Init()
Dim K As Integer, K1 As Integer, X As String

Dim V
param_Init = Null
'$20060626 JPL$ If Not IsNull(paramEdition_Init(Me.lstErr, Me.cmdContext)) Then Exit Function

dirW.PATH = paramEditionSplf_Folder
filW.Pattern = "@*.*"

filW.PATH = paramEditionCorbeille_Folder



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

Private Sub cboFTP_AS400_File_Click()
txtFTP_NT_File = paramNT_Folder & cboFTP_AS400_File & ".txt"

End Sub


Private Sub cmdContext_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Select Case cmdContext.Caption
    Case Is = constcmdRechercher: Me.PopupMenu mnuContext, vbPopupMenuLeftButton
    Case Is = constcmdAbandonner: cmdContext_Quit
End Select

End Sub

Private Sub cmdFTP_AS400_NT_Click()
blnFTP_Get = True
cmdFTP
End Sub

Private Sub cmdFTP_NT_AS400_Click()
blnFTP_Get = False
cmdFTP

End Sub


Private Sub cmdJPL_TEST_Binaire()
Dim K As Integer, K2 As Integer
Dim wAmj As Long, wCours As Double
Dim X1 As String, X2 As String

Me.Enabled = False
Call lstErr_Clear(lstErr, cmdContext, "cmdJPL_Import_DCN : début")
Dim V, xWhere As String, xSQL As String
Dim xZBASTAB0 As typeZBASTAB0
Dim Nb As Long
Screen.MousePointer = vbHourglass



'cmdJPL_Import_DCN
Set rsSab = Nothing
xWhere = " where BASTABNUM = 37 and BASTABARG like 'AUD%' order by bastabarg"
xWhere = " where BASTABNUM = 25  order by bastabarg"
xSQL = "select * from " & paramIBM_Library_SAB & ".ZBASTAB0 " & xWhere
Set rsSab = cnsab.Execute(xSQL)
Nb = 0
Do While Not rsSab.EOF
        V = rsZBASTAB0_GetBuffer(rsSab, xZBASTAB0)
'_____________________________________________________________________________________________
Nb = Nb + 1
'X1 = Mid$(xZBASTAB0.BASTABARG, 4, 4) '37
X1 = Mid$(xZBASTAB0.BASTABARG, 10, 4) '25
wAmj = CLng(convX2P(Mid$(xZBASTAB0.BASTABARG, 10, 4)))
wCours = CDbl(convX2P(Mid$(xZBASTAB0.BASTABDON, 1, 8)))
'Debug.Print wAmj, wCours
    
X2 = convP2X(wAmj, 7)
If X1 <> X2 Then
    MsgBox Nb & "ERR date : " & xZBASTAB0.BASTABARG & wAmj
    Debug.Print "ERR date : "; wAmj
    Debug.Print Asc(Mid$(X1, 1, 1)), Asc(Mid$(X2, 10, 1))
    Debug.Print Asc(Mid$(X1, 2, 1)), Asc(Mid$(X2, 11, 1))
    Debug.Print Asc(Mid$(X1, 3, 1)), Asc(Mid$(X2, 12, 1))
    Debug.Print Asc(Mid$(X1, 4, 1)), Asc(Mid$(X2, 13, 1))
    Debug.Print "---------------------------------"
End If

If wAmj > 1080000 Then

    X1 = Mid$(xZBASTAB0.BASTABDON, 1, 8)
    X2 = convP2X(wCours, 15)
    If X1 <> X2 Then
        MsgBox Nb & "ERR cours : " & xZBASTAB0.BASTABARG & wCours
        Debug.Print "ERR cours : "; wCours
        Debug.Print Asc(Mid$(X1, 1, 1)), Asc(Mid$(X2, 1, 1))
        Debug.Print Asc(Mid$(X1, 2, 1)), Asc(Mid$(X2, 2, 1))
        Debug.Print Asc(Mid$(X1, 3, 1)), Asc(Mid$(X2, 3, 1))
        Debug.Print Asc(Mid$(X1, 4, 1)), Asc(Mid$(X2, 4, 1))
        Debug.Print Asc(Mid$(X1, 5, 1)), Asc(Mid$(X2, 5, 1))
        Debug.Print Asc(Mid$(X1, 6, 1)), Asc(Mid$(X2, 6, 1))
        Debug.Print Asc(Mid$(X1, 7, 1)), Asc(Mid$(X2, 7, 1))
        Debug.Print Asc(Mid$(X1, 8, 1)), Asc(Mid$(X2, 8, 1))
        Debug.Print "---------------------------------"
    End If
End If
'_____________________________________________________________________________________________

   rsSab.MoveNext

Loop
    


Call lstErr_AddItem(lstErr, cmdContext, "cmdJPL_test binaire: fin " & Nb)
Screen.MousePointer = 0
Me.Enabled = True

End Sub

Private Sub cmdJPL_TEST()
Dim K As Integer, K2 As Integer
Dim wAmj As Long, wCours As Double
Dim X1 As String, X2 As String
Dim mDATAVINUM As Long, xDATAVINUM As Long
Dim mDATAVIINT As Currency, xDATAVIINT As Currency
Dim blnOk As Boolean
Me.Enabled = False
Call lstErr_Clear(lstErr, cmdContext, "cmdJPL_Import_DCN : début")
Dim V, xWhere As String, xSQL As String
Dim xZBASTAB0 As typeZBASTAB0
Screen.MousePointer = vbHourglass
Dim cnX As New ADODB.Connection
Dim rsX As New ADODB.Recordset
Dim colPC_User As New Dictionary



        cnX.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        cnX.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" _
        & "SERVER=192.168.168.16;" _
        & "UID=sysinvent_adm;" _
        & "PWD=Manage;" _
        & "OPTION=" & 1 + 2 + 8 + 32 + 2048 + 16384

        cnX.Open
        cnX.Execute ("USE sysinvent") 'select database
        'MsgBox(strQuery)
     '   rs.Open(strQuery, conn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
     '   conn.Close()
X = "select * from system where Sys_Info = 'UserName'"
Set rsX = cnX.Execute(X)
K = 0
Do While Not rsX.EOF
    K = K + 1
    Debug.Print K, rsX("PC"), rsX("Sys_Valeur")
   ' colPC_User.Add rsX("PC"), rsX("Sys_Valeur")
   rsX.MoveNext

Loop

Call lstErr_AddItem(lstErr, cmdContext, "cmdJPL_Import_DCN : fin")
Screen.MousePointer = 0
Me.Enabled = True

Exit Sub


colPC_User.RemoveAll
cnX.Open "DSN=SysInvent"
Set rsX = Nothing
X = "select * from system where Sys_Info = 'UserName'"
Set rsX = cnX.Execute(X)
K = 0
Do While Not rsX.EOF
    K = K + 1
    Debug.Print K, rsX("PC"), rsX("Sys_Valeur")
   ' colPC_User.Add rsX("PC"), rsX("Sys_Valeur")
   rsX.MoveNext

Loop
X1 = colPC_User("0")
Debug.Print X1
X1 = colPC_User.Item("CZC4381JYW")
Debug.Print X1


cnX.Close
Screen.MousePointer = 0
Me.Enabled = True


Exit Sub
'cmdJPL_Import_DCN
Set rsSab = Nothing

xSQL = "select * from SAB073U.ZREPSIT0 "
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    Debug.Print rsSab("REPSITMV") & rsSab("REPSITDAT") & rsSab("REPSITCIB") & rsSab("REPSITLC") & rsSab("REPSITDOC")
'REPSITMV
'REPSITDAT
'REPSITCIB
'REPSITLC
'REPSITDOC
'REPSITFE
'REPSITZA
'REPSITMO
'REPSITFIL xDATAVINUM = rsSab("DATAVINUM")
'_____________________________________________________________________________________________

   rsSab.MoveNext

Loop
    


Call lstErr_AddItem(lstErr, cmdContext, "cmdJPL_Import_DCN : fin")
Screen.MousePointer = 0
Me.Enabled = True

End Sub


Private Sub cmdJPL_click()

Dim arrX2P(256) As String * 2, arrP2X(256) As String * 2
Dim X2P, P2X
Dim K As Integer, K2 As Integer
Dim wCacls As String, wFileName_Splf As String, xUser_Id As String
Dim IdShell, X As String
Dim wUser As typeUser
Dim xSQL As String, Nb As Long

cmdJPL_DAUTLIB0

'cmdJPL_BICBIC
'cmdJPL_SAB_DS
'cmdJPL_TEST_Binaire

'Exit Sub

'________________________________________________________
'Call cmdJPL_DRENTA
Exit Sub

'cmdJPL_Import_ZXSWITG2
'cmdJPL_Import_ZSWITG2
'cmdJPL_Import_ZXBICSRD

'cmdJPL_Import_INSEE
'Dim usr As IADsUser
'Set usr = gestobject("LDAP://CN=test,CN=users,DC=bia-paris,DC=fr")
'Debug.Print usr.FullName

'cmdJPL_Lagarde
'cmdJPL_World_Check
'cmdJPL_Import_DCN
'Set rsSab = Nothing
'xSql = " (MONAPP,MONFLUX)"
'X = "'line1" & vbCr & "line2'"
'xValues = " values('Test2'," & X & ")"

'xSql = "Insert into JPLTST.JPL070504" & xSet & xValues

'Set rsSab = cnsab.Execute(xSql, Nb)

' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    MsgBox "Erreur màj : "
End If
 



xSQL = "select * from JPLTST.JPL070504"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    Debug.Print rsSab("MONAPP"), rsSab("MONFLUX")
   rsSab.MoveNext

Loop

Exit Sub

'wFileName_Splf = "\\BIADOCSRV\.biadoc$\SPLF\test\test.txt"
wFileName_Splf = paramServer("\\BiaDoc\SPLF\test\test.txt")
wUser.Id = "GRACHEH_TC"
Call Table_User(wUser)
xUser_Id = wUser.Id
If Trim(wUser.AliasWin) <> "" Then xUser_Id = wUser.AliasWin
Call File_CACLS(wFileName_Splf, xUser_Id, wUser.Unit)


Exit Sub

Me.Enabled = False
Call lstErr_Clear(lstErr, cmdContext, "cmdJPL_Import_DCN : début")
Dim V, xWhere As String
Dim xZBASTAB0 As typeZBASTAB0
cmdJPL_TEST
Exit Sub

Screen.MousePointer = vbHourglass



'cmdJPL_Import_DCN
Set rsSab = Nothing
xWhere = " where BASTABNUM = 999 and BASTABARG like 'JPL%' order by bastabarg"
xSQL = "select * from " & paramIBM_Library_SAB & ".ZBASTAB0 " & xWhere
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
        V = rsZBASTAB0_GetBuffer(rsSab, xZBASTAB0)
    For K = 0 To 256
        arrX2P(K) = "": arrP2X(K) = ""
    Next K
    
    For K = 0 To 129
        X = Mid$(xZBASTAB0.BASTABDON, K + 1, 1)
        K2 = Asc(X)
        If K = 130 Then
            Debug.Print "' "; K; K2
        End If
        If arrX2P(K2) <> "  " Then MsgBox "X2P " & K & " " & K2 & " " & arrX2P(K2)
        If arrP2X(K) <> "  " Then MsgBox "P2X " & K & " " & K2 & " " & arrP2X(K)
        arrP2X(K) = X
        If K < 100 Then
            arrX2P(K2) = Format$(K, "00")
        Else
            If K < 110 Then
                arrX2P(K2) = Format$(K - 100, "0") & " "
            
             Else
                If K < 120 Then
                    arrX2P(K2) = Format$(K - 110, "0") & " "
                Else
                    arrX2P(K2) = Format$(K - 120, "0") & " "
                End If
            End If
        End If
    Next K
    For K = 0 To 255
        
'        Debug.Print Chr$(34) & arrX2P(K) & Chr$(34) & ",";
        Debug.Print "chr$(" & Asc(arrP2X(K)) & "),";
        If K Mod 10 = 9 Then Debug.Print " _"
        If K Mod 100 = 99 Then Debug.Print

    Next K
      
   rsSab.MoveNext

Loop
    


Call lstErr_AddItem(lstErr, cmdContext, "cmdJPL_Import_DCN : fin")
Screen.MousePointer = 0
Me.Enabled = True

End Sub

Private Sub cmdSplf_Corbeille_Clear_Click()
Dim K As Integer, X8 As String * 8
On Error Resume Next
'On Error GoTo Error_Handle
Me.Enabled = False

Screen.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "cmdSplf_Corbeille_Clear : début")
Nb = 0
Call DTPicker_Control(txtSplfAmj, wSplfAmj)

filW.PATH = Trim(dirW.PATH) & "\Corbeille"
filW.Pattern = "*.*"

filW.Visible = False

For K = 0 To filW.ListCount - 1
    filW.ListIndex = K
    Set fsoFile = msFileSystem.GetFile(filW.PATH & "\" & filW.FileName)
    If Err = 0 Then
        Call dateJMA6_AMJ(fsoFile.DateLastModified, X8)
        If X8 < wSplfAmj Then Nb = Nb + 1: msFileSystem.DeleteFile filW.PATH & "\" & filW.FileName, True
    End If
    If Err > 0 Then Call lstErr_AddItem(lstErr, cmdContext, Err & " : " & filW.PATH & "\" & filW.FileName): Err = 0
    
Next K
filW.PATH = Trim(dirW.PATH)
filW.PATH = Trim(dirW.PATH) & "\Corbeille"
filW.Pattern = "*.*"
filW.Visible = True
Call lstErr_AddItem(lstErr, cmdContext, "cmdSplf_Corbeille_Clear : Fin " & Nb & " fichiers")
Screen.MousePointer = vbDefault
Me.Enabled = True
Exit Sub

Error_Handle:
Shell_MsgBox "#cmdSplf_Corbeille_Clear_Click# " & filW.FileName & ":" & Error, vbCritical, Me.Caption, False
Close
Me.Enabled = True


End Sub

Private Sub cmdSplf_Corbeille_Click()
Dim K1 As Integer, K As Integer, X As String
Dim kSpécial As Integer

Me.Enabled = False
Call lstErr_Clear(lstErr, cmdContext, "cmdSplf_Corbeille : début")
Screen.MousePointer = vbHourglass
Call DTPicker_Control(txtSplfAmj, wSplfAmj)
K1 = Len(Trim(dirW.PATH))
Nb = 0
For K = 0 To dirW.ListCount - 1
    X = dirW.List(K)
    filW.PATH = X   ' paramEditionSplf_Folder & Trim(mId$(X, K1 + 2, Len(X) - K1)) & "\"
    X = Trim(Mid$(X, K1 + 2, Len(X) - K1))
    filW.Pattern = "*.*"
'2005.06.08 ne pas effacer Corbeille et Archive.....
    'Select Case X
    '    Case "Archive", "Corbeille"
    '    Case Else: cmdSPLF_Corbeille_Move
    'End Select
    kSpécial = InStr(X, "Corbeille")
    
    If kSpécial = 0 Then kSpécial = InStr(X, "Archive")
    If kSpécial = 0 Then kSpécial = InStr(X, "NoPaper")
    If kSpécial = 0 Then cmdSPLF_Corbeille_Move
    
Next K
Call lstErr_AddItem(lstErr, cmdContext, "cmdSplf_Corbeille : Fin " & Nb & " fichiers")
Screen.MousePointer = vbDefault
Me.Enabled = True

End Sub

Private Sub cmdSplf_NoPaper_Clear_Click()
Dim K As Integer, X8 As String * 8, X As String
Dim objFolder As Scripting.Folder, objSubFolders As Scripting.Folders

On Error Resume Next
'On Error GoTo Error_Handle
Me.Enabled = False

Screen.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "cmdSplf_NoPaper_Clear : début")

'Suppression des répertoires \NoPaper\DOC antérieurs à 40 jours
'----------------------------------------------------------------
X8 = dateElp("Jour", -40, DSys)

Set objFolder = msFileSystem.GetFolder(paramEditionNoPaper_Folder & "DOC\")
Set objSubFolders = objFolder.SubFolders

For Each objFolder In objSubFolders
    K = InStr(objFolder.Name, "_")
    If K > 0 Then
        X = Mid$(objFolder.Name, K + 1, Len(objFolder.Name) - K)
        If Len(X) = 8 Then
            If X < X8 Then
                Call lstErr_AddItem(lstErr, cmdContext, "delete DOC\" & objFolder.Name)
                msFileSystem.DeleteFolder paramEditionNoPaper_Folder & "DOC\" & objFolder.Name
            End If
        End If
    End If
Next

'Suppression des répertoires \NoPaper\PDF\PROD_ antérieurs à 7 jours
'----------------------------------------------------------------
X8 = dateElp("Jour", -7, DSys)

Set objFolder = msFileSystem.GetFolder(paramEditionNoPaper_Folder & "PDF\")
Set objSubFolders = objFolder.SubFolders

For Each objFolder In objSubFolders
    K = InStr(objFolder.Name, "Prod_")
    If K > 0 Then
        X = Mid$(objFolder.Name, K + 5, Len(objFolder.Name) - K - 4)
        If Len(X) = 8 Then
            If X < X8 Then
                Call lstErr_AddItem(lstErr, cmdContext, "delete PDF\" & objFolder.Name)
                msFileSystem.DeleteFolder paramEditionNoPaper_Folder & "PDF\" & objFolder.Name
            End If
        End If
    End If
Next


Call lstErr_AddItem(lstErr, cmdContext, "cmdSplf_NoPaper_Clear : Fin ")
Screen.MousePointer = vbDefault
Me.Enabled = True
Exit Sub

Error_Handle:
Shell_MsgBox "#cmdSplf_NoPaper_Clear_Click# " & filW.FileName & ":" & Error, vbCritical, Me.Caption, False
Close
Me.Enabled = True

End Sub

Private Sub cmdSplfMonitor_Dir_Click()
Dim wNb_Ftp As Long

If Me.Enabled And blncmdSplfMonitor Then Exit Sub
blncmdSplfMonitor = True
filDoc.Visible = True
Me.Enabled = False

Call lstErr_Clear(lstErr, cmdContext, "cmdSplfMonitor_AS400: " & Time)
Dim I As Integer, K As Integer, X As String, L As Integer, iSession As Integer


filDoc.PATH = paramFTP_SPLF
filDoc.Pattern = "SPLF*.XXX" 'pour réinitialiser
filDoc.Pattern = "SPLF*.txt"
For I = 0 To filDoc.ListCount - 1
    filDoc.ListIndex = I
    txtSplfMonitor_Import = paramFTP_SPLF & "\" & Trim(filDoc.FileName)
    cmdSplfMonitor_Ftp_Click
Next I
filDoc.Pattern = "SPLF*.XXX" 'pour réinitialiser
filDoc.Pattern = "SPLF*.txt"

Me.Enabled = True
blncmdSplfMonitor = False

End Sub
Private Sub cmdSplfMonitor_Ftp_Click()
Dim xIn As String, X As String
Dim seq As Long
Dim wFileName_FTP As String
Dim paramSplfMonitor_Import As String, paramSplfMonitor_Export As String
Dim paramSplfMonitor_Folder As String, paramSplfMonitor_Write As String
Dim wEdition_Id As String
Dim blnHold As Boolean, wPut_Folder As String
Dim wDestinataire As String
Dim xUser As typeUser, xUnit As typeUnit
Dim xEdition_Form As typeEdition_Form
Dim wFileName_Splf As String, wFileName_Archive As String, wFileName_X As String
Dim blnDestinataire_Mod As Boolean, wDestinataire_Mod As String
Dim blnCHGVE053P1  As Boolean
Dim blnSAB_A8 As Boolean
Dim blnSendMail As Boolean, blnNoPaper_Ok As Boolean
Dim wUser_CACLS As typeUser
'20080828
Dim blnSalaires_Check As Boolean, blnSalaires_True As Boolean, kSalaires As Integer

'20050331 Affectation des états SCHGE008P1 par service en fonction du code opération
Dim blnSCHGE008P1 As Boolean, kSCHGE008P1 As Integer, sseSCHGE008P1 As String, opeSCHGE008P1 As String

'$JPL 2014-11-12 états cautions DAFI ou SOBI
Dim blnCaution_Check As Boolean, kCaution As Integer
Dim blnCHGZE_Check As Boolean, kCHGZE As Integer

Dim wFileName_NoPaper As String

On Error GoTo Error_Handle

Me.Enabled = False

blncmdSplfMonitor = True
paramSplfMonitor_Export = ""
blnError = True
Nb_Ftp = 0
lstErr.Visible = True
Call lstErr_Clear(lstErr, cmdContext, "cmdSplfMonitor_Ftp_Click : " & Time)

paramSplfMonitor_Import = Trim(txtSplfMonitor_Import)
wFileName_FTP = Dir(paramSplfMonitor_Import)
If wFileName_FTP = "" Then Call lstErr_AddItem(lstErr, cmdContext, "! pas de fichier : " & paramSplfMonitor_Import): Exit Sub
paramSplfMonitor_Folder = Trim(txtSplfMonitor_Export)
paramSplfMonitor_Export = paramSplfMonitor_Folder & "FTP\" & wFileName_FTP
Call FEU_ROUGE
If Dir(paramSplfMonitor_Export) <> "" Then Kill paramSplfMonitor_Export

msFileSystem.MoveFile paramSplfMonitor_Import, paramSplfMonitor_Export

Open paramSplfMonitor_Export For Input As #1


Do Until EOF(1)
    seq = seq + 1
 '   If Seq Mod 100 = 0 Then Call lstErr_Clear(frmElpKM.lstErr, frmElpKM.cmdContext, "cmdSPLF_Click : " & Seq)
    DoEvents
    Line Input #1, xIn
    If seq <> Val(Mid$(xIn, 1, 9)) Then
        Shell_MsgBox "#cmdSPLFMONITOR_FTP#   seq : " & seq & " <> " & xIn, vbInformation, Me.Caption, True
        seq = Val(Mid$(xIn, 1, 9))
    End If
    If Mid$(xIn, 22, 3) <> "$$$" Then
        If Mid$(xIn, 22, 4) = "$   " Then
            Close 2
            
            blnNoPaper_Ok = True
            
            wEdition_Id = Mid$(xIn, 25 + 20, 10)
            xEdition_Form.K1 = "SAB"
            xEdition_Form.K2 = wEdition_Id
            Call rsEdition_Form(xEdition_Form)
            
            If Mid$(wEdition_Id, 1, 1) = "Q" Then xEdition_Form.Hold = "1": blnNoPaper_Ok = False ' HOLD impression système
            
            xUser.Id = Mid$(xIn, 25 + 30, 10)
            
            
            Call Table_User(xUser)
            ' Ajout Test SAB version le 17.03.2005
            blnSAB_A8 = False
            If Mid$(xIn, 25 + 102, 7) = "0697254" Then                    ' SABA : 65BF511
                blnSAB_A8 = True                                         ' I5A7 : 659807E
                xUser.ProdTest = "T"                                     ' P6A8 : 0697254
            End If
            
            'JPL test   xUser.QSYSOPR = "1"

            wDestinataire = xUser.Id
            blnCHGVE053P1 = False: blnDestinataire_Mod = False
            '20050331
            blnSCHGE008P1 = False
            blnCaution_Check = False
            blnSalaires_Check = False: blnSalaires_True = False
            
            If Trim(wEdition_Id) = "SITTE014P1" Then blnSalaires_Check = True
            If Trim(wEdition_Id) = "SITTE028P1" Then blnSalaires_Check = True
            If Trim(wEdition_Id) = "SITTE043P1" Then blnSalaires_Check = True
            If InStr(wEdition_Id, "CAUT") > 0 Then blnCaution_Check = True: kCaution = 0
            If InStr(wEdition_Id, "CHGZE") > 0 Then blnCHGZE_Check = True: kCHGZE = 0
            If InStr(wEdition_Id, "SCA601") > 0 Then blnCaution_Check = True: kCaution = 0
                
            If xUser.QSYSOPR = "1" Then
            
                If Trim(wEdition_Id) = "AUT329P3" Then xEdition_Form.Hold = "1"
                If Trim(wEdition_Id) = "CPT051P1" Then xEdition_Form.Hold = "1"
                If Trim(wEdition_Id) = "CHGVE053P1" Then blnCHGVE053P1 = True
                ' 20050331
                If Trim(wEdition_Id) = "SCHGE008P1" Then blnSCHGE008P1 = True: kSCHGE008P1 = 0: opeSCHGE008P1 = "": sseSCHGE008P1 = ""
                

                If Trim(xEdition_Form.Unit) <> "" Then
                    If xUser.ProdTest = "T" Then
                        wDestinataire = "_T_" & xEdition_Form.Unit
                    Else
                        wDestinataire = "_" & xEdition_Form.Unit
                    End If
                End If
            End If
            
            Select Case xUser.ProdTest
                Case "P"
                            If paramEdition_Print_Production _
                            And xUser.Edition_Hold = "0" _
                            And xEdition_Form.Hold = "0" Then
                                wPut_Folder = "Print\"
                            Else
                                wPut_Folder = "Production\"
                            End If
                Case "T"
                            If paramEdition_Print_Test _
                            And xUser.Edition_Hold = "0" _
                            And xEdition_Form.Hold = "0" Then
                                wPut_Folder = "Print\"
                            Else
                                wPut_Folder = "Test\"
                            End If
                Case "I": wPut_Folder = "System\"
                Case Else:
                        If Mid$(xUser.Id, 1, 2) = "T_" Or Mid$(xUser.Id, 1, 2) = "V_" Then
                            blnNoPaper_Ok = False
                            wPut_Folder = "Test\"
                        Else
                            If Mid$(xUser.Id, 1, 2) = "I_" Then
                                blnNoPaper_Ok = False
                                 wPut_Folder = "System\"
                            Else
                                wPut_Folder = "Production\"
                            End If
                       End If
            End Select
            
            wDestinataire_Mod = Trim(wDestinataire)
            wFileName_X = "." & Mid$(xIn, 25 + 1, 8) & "_" & Mid$(xIn, 25 + 61, 6) & "_" & Trim(Mid$(xIn, 25 + 20, 10)) & "_" & Mid$(xIn, 25 + 9, 6) & "_" & Mid$(xIn, 25 + 15, 5) & ".txt"

            ' Test version A8 le 17.03/2005
            If blnSAB_A8 Then
                wFileName_X = "." & Mid$(xIn, 25 + 1, 8) & "_" & Mid$(xIn, 25 + 61, 6) & "_" & Trim(Mid$(xIn, 25 + 20, 10)) & "_A8-" & Mid$(xIn, 25 + 9, 6) & "_" & Mid$(xIn, 25 + 15, 5) & ".txt"
            End If
            
            wFileName_Splf = paramEditionSplf_Folder & wPut_Folder & wDestinataire_Mod & wFileName_X
            
            Call lstErr_ChangeLastItem(lstErr, cmdContext, X)
            Call FEU_ROUGE
            Open wFileName_Splf For Output As #2
            Nb_Ftp = Nb_Ftp + 1
        End If
        
        Print #2, Mid$(xIn, 22, Len(xIn) - 21)
       
        If blnSalaires_Check Then
            If Not blnSalaires_True Then
                kSalaires = InStr(Mid$(xIn, 22, 25), "DRH")
                

                If kSalaires > 0 Then
                    blnSalaires_True = True
                    blnSalaires_Check = False
                    'blnDestinataire_Mod = True
                    'wDestinataire_Mod = "_DRH"
                    'wPut_Folder = "Production\"
                End If
            End If
       End If
            

        
        If blnCHGVE053P1 Then
            Debug.Print Mid$(xIn, 22, 18)
            If Mid$(xIn, 22, 18) = "   1    Nos réf. :" Then
                blnCHGVE053P1 = False
                blnDestinataire_Mod = True
                If IsNumeric(Mid$(xIn, 47, 3)) Then
                    wDestinataire_Mod = "_SOBF"
                Else
                     wDestinataire_Mod = "_ORPA"
               End If
            End If
        End If
        
        ' 20050331
        If blnSCHGE008P1 Then
            If kSCHGE008P1 <= 0 Then
                If sseSCHGE008P1 = "" Then
                    kSCHGE008P1 = InStr(xIn, "SCHGE008P1/")
                    If kSCHGE008P1 > 0 Then sseSCHGE008P1 = Mid$(xIn, kSCHGE008P1 + 11, 4)
                End If
                kSCHGE008P1 = InStr(xIn, "OPE/EVE/N")                ' position du texte
            Else
                If opeSCHGE008P1 = "" Then
                    opeSCHGE008P1 = Trim(Mid$(xIn, kSCHGE008P1, 3))  ' Code opération
                    If opeSCHGE008P1 <> "" Then
                        wDestinataire_Mod = "_" & Table_Ope_Unit(sseSCHGE008P1 & opeSCHGE008P1) ' Service
                        
                        blnDestinataire_Mod = True
                        blnSCHGE008P1 = False                        ' Fin de la recherche
                    End If
                End If
            End If
        End If
        
'$JPL 2014-11-12 états cautions DAFI ou SOBI
'________________________________________________________________________
        If blnCaution_Check Then
            kCaution = kCaution + 1
            If kCaution > 1 Then
                blnCaution_Check = False
                If InStr(xIn, "/00CD") > 0 Then
                    blnDestinataire_Mod = True
                    wDestinataire_Mod = "_SOBI"
                End If
             End If
       End If
        If blnCHGZE_Check Then
            kCHGZE = kCHGZE + 1
            If kCHGZE > 1 Then
                blnCHGZE_Check = False
                If InStr(xIn, "/00CD") > 0 Then
                    blnDestinataire_Mod = True
                    wDestinataire_Mod = "_SOBI"
                End If
             End If
       End If
'________________________________________________________________________
        If Mid$(xIn, 22, 4) = "$$  " Then
            meSplfJob.SJQXAMJ = Mid$(xIn, 25 + 87, 8)
            meSplfJob.SJQXHMS = Mid$(xIn, 25 + 95, 8)
            Close 2
            
            wUser_CACLS.Id = wDestinataire_Mod
            Call Table_User_CACLS(wUser_CACLS)
            Call File_CACLS(wFileName_Splf, wUser_CACLS.Id, wUser_CACLS.Unit)    'gestion des droits  ACL

            If xUser.ProdTest = "P" Then
'________________________________________________________________________________________________
'                Call File_CACLS(wFileName_Splf, wUser_Id, wUser_Unit)     'gestion des droits  ACL
'__________________________________________________________________________________________________
                blnSendMail = False
                If InStr(wFileName_X, "SCHGE005P1") > 0 Then blnSendMail = True
                If InStr(wFileName_X, "FCIGS018P1") > 0 Then blnSendMail = True
                If InStr(wFileName_X, "FCIGS018P3") > 0 Then blnSendMail = True
                If blnSendMail Then
                    wFileName_Archive = paramEditionSplf_Folder & "SendMail\" & wDestinataire_Mod & wFileName_X
                    msFileSystem.CopyFile wFileName_Splf, wFileName_Archive
                End If
            End If
            If xEdition_Form.Save = "1" Then
                wFileName_Archive = paramEditionSplf_Folder & "Archive\" & wDestinataire_Mod & wFileName_X
                msFileSystem.CopyFile wFileName_Splf, wFileName_Archive
            End If
            
           If blnDestinataire_Mod Then
                X = paramEditionSplf_Folder & wPut_Folder & wDestinataire_Mod & wFileName_X
                If wFileName_Splf <> X Then
                    msFileSystem.MoveFile wFileName_Splf, X
                    wFileName_Splf = X
                End If
            End If
            
           If blnSalaires_True Then
                msFileSystem.DeleteFile wFileName_Splf
            Else
                If blnNoPaper_Ok Then
                    wFileName_NoPaper = paramEditionNoPaper_Folder & "TXT\" & wDestinataire_Mod & wFileName_X
                    msFileSystem.CopyFile wFileName_Splf, wFileName_NoPaper
                End If
                
                
           End If
           

           

        End If
   End If
Loop
Close
Call FEU_VERT

blnError = False

Call lstErr_Clear(lstErr, cmdContext, "cmdSPLF_Click fin : " & seq)
Me.Enabled = True

Exit Sub

Error_Handle:
Close
Shell_MsgBox "#cmdSplfMonitor_Ftp_Click# " & paramSplfMonitor_Export & " : " & Error, vbCritical, Me.Caption, False
Me.Enabled = True
End Sub




Private Sub cmdZXSWITG2_Click()
cmdJPL_Import_ZXSWITG2
End Sub

Private Sub fgSelect_Click()
fgSelect.LeftCol = 0

End Sub

Private Sub fgSelect_LeaveCell()
On Error Resume Next
fgSelect.CellBackColor = &HE0E0E0
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
SendKeys "{TAB}"
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
Dim wRow As Long
On Error Resume Next
If y <= fgSelect.RowHeightMin Then
    Select Case fgSelect.Col
        Case 0: fgSelect_Sort1 = 0: fgSelect_Sort2 = 2: fgSelect_Sort
        Case 1:  fgSelect_SortX 1
        Case 2: fgSelect_SortX 2
        Case 3: fgSelect_Sort1 = 3: fgSelect_Sort2 = 3: fgSelect_Sort
        Case fgSelect_arrIndex:  fgSelect_SortX fgSelect_arrIndex
    End Select
Else
    If fgSelect.Rows > 1 Then
        If fgSelect.Col = 2 Or fgSelect.Col = 1 Then
            fgSelect_RowClick = fgSelect.Row
            fgSelect_ColClick = fgSelect.Col
            fgSelect.CellBackColor = vbCyan
       End If
    '   fgSelect_MouseDown_Ok
   ' If Button = vbRightButton Then
     '   Me.PopupMenu mnuSelect, vbPopupMenuLeftButton
   End If
End If
End Sub

Public Sub fgSelect_Reset()
fgSelect.Clear
fgSelect_Sort1 = 0: fgSelect_Sort2 = 0
fgSelect_Sort1_Old = -1
fgSelect_RowDisplay = 0: fgSelect_RowClick = 0
fgSelect_arrIndex = 6
blnfgSelect_DisplayLine = False
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

End Sub


Public Sub Auto_SplfJob()
Dim wMsg As String * 100
If blnAuto_SplfJob_Run = False Then
    blnAuto_SplfJob_Run = True
    
    cmdSplfMonitor_Dir_Click  '$ 2003.09.02 jpl   cmdSplfMonitor_AS400_Click
    
    
    wMsg = "@PRINT_PROD"
    frmEdition.Msg_Rcv wMsg:
    blnAuto_SplfJob_Run = False
End If

Unload Me

End Sub

Private Sub cboFTP_AS400_File_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub






Private Sub txtSplfAmj_GotFocus()
DTPicker_GotFocus txtSplfAmj

End Sub


Private Sub txtSplfAmj_LostFocus()
DTPicker_LostFocus txtSplfAmj

End Sub



Public Sub cmdSPLF_Corbeille_Move()
Dim K As Integer, xFileName As String, X8 As String * 8
On Error Resume Next
filW.Visible = False
Call FEU_ROUGE
For K = 0 To filW.ListCount - 1
    filW.ListIndex = K
    xFileName = filW.PATH & "\" & filW.FileName
    Set fsoFile = msFileSystem.GetFile(filW.PATH & "\" & filW.FileName)
    If Err = 0 Then
        Call dateJMA6_AMJ(fsoFile.DateLastModified, X8)
        If X8 < wSplfAmj Then Nb = Nb + 1: msFileSystem.MoveFile xFileName, paramEditionSplf_Folder & "Corbeille\" & filW.FileName
    End If
    If Err > 0 Then Call lstErr_AddItem(lstErr, cmdContext, Err & " : " & filW.PATH & "\" & filW.FileName): Err = 0
Next K
Call FEU_VERT
filW.Visible = True
End Sub

Public Sub cmdFTP()
Dim wFTP_Nt_Filename As String, wFTP_Nt_Filename_Bat As String, wFTP_Nt_Filename_Dta As String
Dim blnFTP_AS400_Binary As Boolean

On Error GoTo Error_Handle

Me.Enabled = False

wFTP_Nt_Filename = Trim(txtFTP_NT_File)
Call lstErr_Clear(Me.lstErr, Me.cmdContext, "cmdFTP : " & Time)
If wFTP_Nt_Filename = "" Then Call lstErr_AddItem(Me.lstErr, Me.cmdContext, "? préciser le fichier NT ")
If Mid$(wFTP_Nt_Filename, Len(wFTP_Nt_Filename), 1) = "\" Then Call lstErr_AddItem(Me.lstErr, Me.cmdContext, "? nom de fichier NT Invalide")
If Trim(cboFTP_AS400_Library) = "" Then Call lstErr_AddItem(Me.lstErr, Me.cmdContext, "? préciser la librairie AS400 ")
If Trim(cboFTP_AS400_File) = "" Then Call lstErr_AddItem(Me.lstErr, Me.cmdContext, "? préciser le fichier AS400 ")

If chkFTP_AS400_Binary = "1" Then
    blnFTP_AS400_Binary = True
Else
    blnFTP_AS400_Binary = False
End If
paramIBM_AS400_FTP = cboFTP_AS400_host.Text

Call Shell_FTP(wFTP_Nt_Filename, Trim(cboFTP_AS400_Library), Trim(cboFTP_AS400_File), blnFTP_Get, blnFTP_AS400_Binary)
Me.Enabled = True
Exit Sub
Error_Handle:
Shell_MsgBox "#CmdFTP# :" & Error, vbCritical, Me.Caption, False
End Sub

Public Sub Auto_CLROUTQ()
Dim K As Integer, xFileName As String, X8 As String * 8
Dim blnOk As Boolean

On Error Resume Next

xFileName = paramEditionCorbeille_Folder & "@Auto_CLROUTQ"
Set fsoFile = msFileSystem.GetFile(xFileName)
blnOk = True
If Err = 0 Then
    Call dateJMA6_AMJ(fsoFile.DateLastModified, X8)
    If X8 >= DSys Then blnOk = False
   
End If
If blnOk Then

    Call DTPicker_Set(txtSplfAmj, DSys_VeilleOAP)
    '$2003.09.03 jpl cmdAS400Outq_Corbeille_Clear_Click
    cmdSplf_Corbeille_Clear_Click
    
    Call DTPicker_Set(txtSplfAmj, DSys)
    '$2003.09.03 jpl cmdAS400Outq_Corbeille_Click
    cmdSplf_Corbeille_Click
    
    cmdSplf_NoPaper_Clear_Click '$JPL 20141119
    
    Call FEU_ROUGE
    Open xFileName For Output As #1
    Close
    Call FEU_VERT
End If

End Sub

Public Sub cmdJPL_Import_DCN()
Dim X As String, K As Integer, Nb As Long, NbOk As Long
Dim xIn As String
Dim mAmj7 As Long
Dim blnSelect As Boolean
Dim xZMOUVEA0 As typeZMOUVEA0
Dim V
Dim xSQL As String
Dim xWhere As String, xSet As String, xValues As String
Dim X1 As String, X2 As Long, X3 As Currency
On Error GoTo Error_Handler

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Exit Sub
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

Nb = 0: NbOk = 0
Open Trim("C:\Temp\20060207 R911329.csv") For Input As #3
'cnSAB_Transaction "BeginTrans"

Do Until EOF(3)
    Line Input #3, xIn
    Nb = Nb + 1
    K = 0
    X1 = CSV_Scan(xIn, K)
    X2 = CLng(CSV_Scan(xIn, K))
    X3 = CCur(Val(CSV_Scan(xIn, K)))
    xSet = " (MOUVEMCOM,MOUVEMNUM,MOUVEMMON)"
    xValues = " values('" & X1 & "','" & X2 & "'," & cur_P(X3) & ")"
    Call FEU_ROUGE
    xSQL = "Insert into " & paramIBM_Library_SABSPE & ".YCDOR911" & xSet & xValues

    Set rsSab = cnsab.Execute(xSQL, Nb)
    Call FEU_VERT
    
' Tester si la mise à jour a été effectuée
'===================================================================================

    If Nb = 0 Then
        MsgBox "Erreur màj : "
    End If
 

    DoEvents
Loop

'cnSAB_Transaction "Commit"

Close
Call lstErr_AddItem(lstErr, cmdPrint, "cmdJPL_Import_DCN: " & NbOk & "/" & Nb): DoEvents

Exit Sub
'---------------------------------------------------------
Error_Handler:
'---------------------------------------------------------

Close
Call MsgBox(Err & " : " & Error(Err), vbCritical, "cmdJPL_Import_DCN")

End Sub
Public Sub cmdJPL_Import_INSEE()
Dim X As String, K As Integer, Nb As Long, NbOk As Long
Dim xIn As String
Dim mAmj7 As Long
Dim blnSelect As Boolean
Dim xZMOUVEA0 As typeZMOUVEA0
Dim V
Dim xSQL As String
Dim xWhere As String, xSet As String, xValues As String
Dim X1 As String * 1, X2 As String * 1, X3 As String * 2, X4 As String * 3, X5 As String * 3
Dim X6 As String * 1, X7 As String * 2, X8 As String * 1, X9 As String * 5, X10 As String * 70
Dim X11 As String * 5, X12 As String * 70
Dim wIdFile_Destination As Integer

On Error GoTo Error_Handler


Nb = 0: NbOk = 0
Open Trim("C:\Temp\BIA\comsimp2007\comsimp2007.csv") For Input As #3
Call lstErr_Clear(lstErr, cmdContext, "Export : YBASLIN0 "): DoEvents

V = File_Export_Monitor("Output", wIdFile_Destination, "C:\Temp\BIA\comsimp2007\YBASLIN0")
If Not IsNull(V) Then Exit Sub

Do Until EOF(3)
    Line Input #3, xIn
    Nb = Nb + 1
    K = 0
    X1 = Format(CInt(CSV_Scan(xIn, K)), "0")
    X2 = Format(CInt(CSV_Scan(xIn, K)), "0")
    X3 = Format(CInt(CSV_Scan(xIn, K)), "00")
    X4 = CSV_Scan(xIn, K)
    X5 = Format(CInt(CSV_Scan(xIn, K)), "000")
    X6 = Format(CInt(CSV_Scan(xIn, K)), "0")
    X7 = Format(CInt(CSV_Scan(xIn, K)), "00")
    X8 = Format(CInt(CSV_Scan(xIn, K)), "0")
    X9 = CSV_Scan(xIn, K)
    X10 = CSV_Scan(xIn, K)
    X11 = CSV_Scan(xIn, K)
    X12 = CSV_Scan(xIn, K)
    Print #wIdFile_Destination, X1; X2; X3; X4; X5; X6; X7; X8; X9; X10; X11; X12
    
    NbOk = NbOk + 1
   ' If NbOk Mod 1000 Then Call lstErr_ChangeLastItem(lstErr, cmdPrint, "cmdJPL_Import_INSEE: " & NbOk & "/" & Nb): DoEvents
    DoEvents
Loop


Close
Call lstErr_AddItem(lstErr, cmdPrint, "cmdJPL_Import_INSEE: " & NbOk & "/" & Nb): DoEvents

Exit Sub
'---------------------------------------------------------
Error_Handler:
'---------------------------------------------------------

Close

End Sub

Public Sub cmdJPL_Import_ZXBICSRD()
Dim X As String, K As Integer, Nb As Long, NbOk As Long
Dim xIn As String
Dim mAmj7 As Long
Dim blnSelect As Boolean
Dim xZMOUVEA0 As typeZMOUVEA0
Dim V
Dim xSQL As String
Dim xWhere As String, xSet As String, xValues As String
Dim X1 As String * 2
Dim X2 As String * 1
Dim X3 As String * 8
Dim X4 As String * 8
Dim X5 As String * 3
Dim X6 As String * 105
Dim X7 As String * 35
Dim X8 As String * 2
Dim X9 As String * 8
Dim X10 As String * 8
Dim X11 As String * 1
Dim X12 As String * 8
Dim X13 As String * 11
Dim X14 As String * 1
Dim X15 As String * 1
Dim X16 As String * 11
Dim X17 As String * 8
Dim X18 As String * 8
Dim X19 As String * 8

Dim wIdFile_Destination As Integer

On Error GoTo Error_Handler


Nb = 0: NbOk = 0
Open Trim("C:\Temp\sab\ZXBICSRD.csv") For Input As #3
Call lstErr_Clear(lstErr, cmdContext, "Export : ZXBICSRD "): DoEvents

V = File_Export_Monitor("Output", wIdFile_Destination, "C:\Temp\sab\ZXBICSRD")
If Not IsNull(V) Then Exit Sub

Do Until EOF(3)
    Line Input #3, xIn
    Nb = Nb + 1
    K = 0
    X1 = CSV_Scan(xIn, K)
    X2 = CSV_Scan(xIn, K)
    X3 = CSV_Scan(xIn, K)
    X4 = CSV_Scan(xIn, K)
    X5 = CSV_Scan(xIn, K)
    X6 = CSV_Scan(xIn, K)
    X7 = CSV_Scan(xIn, K)
    X8 = CSV_Scan(xIn, K)
    X9 = CSV_Scan(xIn, K)
    X10 = CSV_Scan(xIn, K)
    X11 = CSV_Scan(xIn, K)
    X12 = CSV_Scan(xIn, K)
    X13 = CSV_Scan(xIn, K)
    X14 = CSV_Scan(xIn, K)
    X15 = CSV_Scan(xIn, K)
    X16 = CSV_Scan(xIn, K)
    X17 = CSV_Scan(xIn, K)
    X18 = CSV_Scan(xIn, K)
    X19 = CSV_Scan(xIn, K)

    Print #wIdFile_Destination, X1; X2; X3; X4; X5; X6; X7; X8; X9; X10; X11; X12; X13; X14; X15; X16; X17; X18; X19
    
    NbOk = NbOk + 1
   ' If NbOk Mod 1000 Then Call lstErr_ChangeLastItem(lstErr, cmdPrint, "cmdJPL_Import_INSEE: " & NbOk & "/" & Nb): DoEvents
    DoEvents
Loop


Close
Call lstErr_AddItem(lstErr, cmdPrint, "cmdJPL_Import_ZXBICSRD: " & NbOk & "/" & Nb): DoEvents

Exit Sub
'---------------------------------------------------------
Error_Handler:
'---------------------------------------------------------

Close

End Sub

Public Sub cmdJPL_DRENTA()
Dim X As String, K As Integer, Nb As Long, NbOk As Long, nbR As Integer
Dim xIn As String
Dim mAmj7 As Long
Dim blnSelect As Boolean
Dim xZMOUVEA0 As typeZMOUVEA0
Dim V
Dim xSQL As String
Dim xWhere As String, xSet As String, xValues As String
Dim X1 As String
Dim X2 As String
Dim X3 As String
Dim X4 As String
Dim X5 As String
Dim X6 As Currency
Dim X7 As Currency

'Dim newDRENTA As typeDRENTA, oldDRENTA As typeDRENTA

On Error GoTo Error_Handler



Set rsSab = Nothing

Nb = 0: NbOk = 0
Open Trim("C:\Temp\delalande\20080827_DWH.csv") For Input As #3
Call lstErr_Clear(lstErr, cmdContext, "Export : 20080827_DWH "): DoEvents


Do Until EOF(3)
    Line Input #3, xIn
    Nb = Nb + 1
    K = 0
    X1 = CSV_Scan(xIn, K)
    X2 = CSV_Scan(xIn, K)
    X3 = CSV_Scan(xIn, K)
    X4 = CSV_Scan(xIn, K)
    X5 = CCur(Val(CSV_Scan(xIn, K)))
    X6 = CCur(Val(CSV_Scan(xIn, K)))
    
'    xWhere = " where DRTAVER = 1 " _
             & " and DRTAPER = " & Mid$(X1, 7, 4) & Mid$(X1, 4, 2) & Mid$(X1, 1, 2) _
             & " and DRTAETA = '01' " _
             & " and DRTACLIA = ' ' " _
             & " and DRTACLIB =" & X3 _
             & " and DRTACRTa = " & X2
    
'    xSql = "select * from BODWH.DRENTA " & xWhere
'    Set rsSab = cnsab.Execute(xSql)
'    nbR = 0
'    Do While Not rsSab.EOF
'        V = srvDRENTA_GetBuffer_ODBC(rsSab, oldDRENTA)
'        If Not IsNull(V) Then
'            MsgBox xWhere, vbCritical, "enregistrement non trouvé)"
'        Else
'            nbR = nbR + 1
'            If oldDRENTA.DRTAMMRB <> X5 Then
'                MsgBox xWhere, vbCritical, "oldDRENTA.DRTAMMRB différent"
'           End If
'        End If
'
'        newDRENTA = oldDRENTA
'        newDRENTA.DRTAMMRB = X6
'       V = sqlDRENTA_Update(newDRENTA, oldDRENTA, cnsab)
'        If Not IsNull(V) Then
'                MsgBox xWhere, vbCritical, "newDRENTA.DRTAMMRB erreur"
'        Else
'            NbOk = NbOk + 1
'        End If
'       ' If NbOk Mod 1000 Then Call lstErr_ChangeLastItem(lstErr, cmdPrint, "cmdJPL_Import_INSEE: " & NbOk & "/" & Nb): DoEvents
'        DoEvents
'        rsSab.MoveNext

    Loop
    If nbR <> 1 Then MsgBox xWhere, vbCritical, "nombre réponse : & nbr"

'Loop


Call lstErr_AddItem(lstErr, cmdPrint, "cmdJPL_DRENTA: " & NbOk & "/" & Nb): DoEvents

Set rsSab = Nothing

Close

Exit Sub
'---------------------------------------------------------
Error_Handler:
'---------------------------------------------------------

Close

End Sub
Public Sub cmdJPL_DAUTLIB0()
Dim X As String, K As Integer, Nb As Long, NbOk As Long, nbR As Integer
Dim xIn As String
Dim mAmj7 As Long
Dim blnSelect As Boolean
Dim xZMOUVEA0 As typeZMOUVEA0
Dim V
Dim xSQL As String
Dim xWhere As String, xSet As String, xValues As String
Dim X1 As String
Dim X2 As String
Dim X3 As String
Dim X4 As String
Dim X5 As String
Dim X6 As Currency
Dim X7 As Currency

'Dim newDAUTLIB0 As typeDAUTLIB0

On Error GoTo Error_Handler

cnSab_Update.Open paramODBC_DSN_SAB


Set rsSab = Nothing

Nb = 0: NbOk = 0
Open Trim("C:\Temp\delalande\dautlib0.csv") For Input As #3
Call lstErr_Clear(lstErr, cmdContext, "dautlib0.csv "): DoEvents


Do Until EOF(3)
    Line Input #3, xIn
    Nb = Nb + 1
    K = 0
'    newDAUTLIB0.DAUTLIBCOD = CSV_Scan(xIn, K)
'    newDAUTLIB0.DAUTLIBTXT = CSV_Scan(xIn, K)
'    newDAUTLIB0.DAUTLIBRGP = CSV_Scan(xIn, K)
'    newDAUTLIB0.DAUTLIBELM = CSV_Scan(xIn, K)
'    newDAUTLIB0.DAUTLIBAMO = CSV_Scan(xIn, K)
'    sqlDAUTLIB0_Insert newDAUTLIB0
Loop
    If nbR <> 1 Then MsgBox xWhere, vbCritical, "nombre réponse : & nbr"

'Loop


Call lstErr_AddItem(lstErr, cmdPrint, "cmdJPL_DRENTA: " & NbOk & "/" & Nb): DoEvents

Set rsSab = Nothing
cnSab_Update.Close

Close

Exit Sub
'---------------------------------------------------------
Error_Handler:
'---------------------------------------------------------

Close

End Sub


Public Sub cmdJPL_Import_ZXSWITG2()
Dim X As String, K As Integer, Nb As Long, NbOk As Long
Dim xIn As String
Dim mAmj7 As Long
Dim blnSelect As Boolean
Dim xZMOUVEA0 As typeZMOUVEA0
Dim V
Dim xSQL As String
Dim xWhere As String, xSet As String, xValues As String
Dim X1 As String * 11
Dim X2 As String * 11
Dim X3 As String * 11
Dim X4 As String * 105
Dim X5 As String * 35
Dim X6 As String * 15
Dim X7 As String * 1
Dim X8 As String * 1
Dim X9 As String * 8
Dim X10 As String * 8
Dim X11 As String * 2

Dim wIdFile_Destination As Integer

On Error GoTo Error_Handler


Nb = 0: NbOk = 0
Open Trim("C:\Temp\sab\ZXSWITG2.csv") For Input As #3
Call lstErr_Clear(lstErr, cmdContext, "Export : ZSWITG2 "): DoEvents

V = File_Export_Monitor("Output", wIdFile_Destination, "C:\Temp\sab\ZXSWITG2.txt")
If Not IsNull(V) Then Exit Sub

Do Until EOF(3)
    Line Input #3, xIn
    Nb = Nb + 1
    K = 0
    X1 = CSV_Scan(xIn, K)
    X2 = CSV_Scan(xIn, K)
    X3 = CSV_Scan(xIn, K)
    X4 = CSV_Scan(xIn, K)
    X5 = CSV_Scan(xIn, K)
    X6 = CSV_Scan(xIn, K)
    X7 = CSV_Scan(xIn, K)
    X8 = CSV_Scan(xIn, K)
    X9 = CSV_Scan(xIn, K)
    X10 = CSV_Scan(xIn, K)
    X11 = CSV_Scan(xIn, K)

    Print #wIdFile_Destination, X1; X2; X3; X4; X5; X6; X7; X8; X9; X10; X11
    
    NbOk = NbOk + 1
   ' If NbOk > 10 Then Exit Do
    If NbOk Mod 1000 Then Call lstErr_ChangeLastItem(lstErr, cmdPrint, "ZXSWITG2: " & NbOk & "/" & Nb): DoEvents
    DoEvents
Loop


Close
Call lstErr_AddItem(lstErr, cmdPrint, "cmdJPL_Import_ZSWITG2: " & NbOk & "/" & Nb): DoEvents

Exit Sub
'---------------------------------------------------------
Error_Handler:
'---------------------------------------------------------

Close

End Sub

Public Sub cmdJPL_Lagarde()
Dim X As String, K As Integer, Nb As Long, NbOk As Long
Dim xIn As String
Dim mAmj7 As Long
Dim blnSelect As Boolean
Dim xZMOUVEA0 As typeZMOUVEA0
Dim V
Dim xSQL As String
Dim xWhere As String, xSet As String, xValues As String
Dim X1 As String, X2 As Long, X3 As Currency
On Error GoTo Error_Handler


xSQL = "select * from " & paramIBM_Library_SAB & ".ZAFGENC0"
Set rsSab = cnsab.Execute(xSQL)

Do Until rsSab.EOF

    Debug.Print rsSab("AFGENCOPE"); rsSab("AFGENCCLI")
Loop
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Exit Sub
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

Nb = 0: NbOk = 0
Open Trim("C:\Temp\bia_banque.csv") For Input As #3
Call FEU_ROUGE
Open Trim("C:\Temp\bia_banque_BIC.csv") For Output As #4

Do Until EOF(3)
    Line Input #3, xIn
    Nb = Nb + 1
    K = 0
    X1 = CSV_Scan(xIn, K)

If paramIBM_AS400_ID = "I5A7" Then
    xSQL = "select ADRESSRA1 from " & paramIBM_Library_SAB & ".ZADRESS0" _
         & " where ADRESSTYP= '4' and ADRESSNUM = ' 00" & X1 & "'"
    Set rsSab = cnsab.Execute(xSQL)
    If Not rsSab.EOF Then
        Print #4, Mid$(rsSab("ADRESSRA1"), 11, 11) & ";" & xIn
    Else
        Print #4, "???????????" & ";" & xIn
    End If
Else
    xSQL = "select ADRESSRA12 from " & paramIBM_Library_SAB & ".ZADRESS0" _
         & " where ADRESSTYP= '4' and ADRESSNUM = ' 00" & X1 & "'"
    Set rsSab = cnsab.Execute(xSQL)
    If Not rsSab.EOF Then
        Print #4, rsSab("ADRESSRA12") & ";" & xIn
    Else
        Print #4, "???????????" & ";" & xIn
    End If
End If
    

Loop


Close
Call FEU_VERT
Call lstErr_AddItem(lstErr, cmdPrint, "cmdJPL_Import_DCN: " & NbOk & "/" & Nb): DoEvents

Exit Sub
'---------------------------------------------------------
Error_Handler:
'---------------------------------------------------------

Close
Call MsgBox(Err & " : " & Error(Err), vbCritical, "cmdJPL_Import_DCN")


End Sub





Public Sub cmdJPL_SAB_DS()
Dim X As String, K As Integer, Nb As Long, NbOk As Long
Dim xIn As String
Dim arrProduit(1000) As String, arrDoc_Titre(1000) As String, arrDoc_Nom(1000) As String, arrDoc_Id(1000) As String
Dim X1 As String, X2 As String, X3 As String, X4 As String
Dim wIdFile_Destination As Integer
Dim K1 As Integer, K2 As Integer
Dim wType As String
Dim blnOk As Boolean
On Error GoTo Error_Handler


Nb = 0: NbOk = 0
Open Trim("C:\Temp\sab_ref.csv") For Input As #3
Call lstErr_Clear(lstErr, cmdContext, "cmdJPL_SAB_DS : C:\Temp\sab_ref.csv"): DoEvents

V = File_Export_Monitor("Output", wIdFile_Destination, "C:\Temp\SAB_Index.htm")
If Not IsNull(V) Then Exit Sub
wType = MsgBox("(doc,titre,produit) = 'OUI'", vbYesNo, "choix du fichier .csv")
If wType = vbYes Then
    Do Until EOF(3)
        Line Input #3, xIn
        Nb = Nb + 1
        K = 0
        arrDoc_Nom(Nb) = CSV_Scan(xIn, K)
        arrDoc_Titre(Nb) = CSV_Scan(xIn, K)
        arrProduit(Nb) = CSV_Scan(xIn, K)
        
        DoEvents
    Loop

Else
    Do Until EOF(3)
        Line Input #3, xIn
        Nb = Nb + 1
        K = 0
        arrProduit(Nb) = CSV_Scan(xIn, K)
        arrDoc_Titre(Nb) = CSV_Scan(xIn, K)
        X = CSV_Scan(xIn, K)
        X = CSV_Scan(xIn, K)
        X = CSV_Scan(xIn, K)
        arrDoc_Nom(Nb) = CSV_Scan(xIn, K)
        
        DoEvents
    Loop
End If
Open Trim("C:\Temp\sab_ds.txt") For Input As #4
Call lstErr_AddItem(lstErr, cmdContext, "cmdJPL_SAB_DS : C:\Temp\sab_ds.txt"): DoEvents

Do Until EOF(4)
    Line Input #4, xIn
    blnOk = False
    K = 0
    X = CSV_Scan(xIn, K)
    X4 = Mid$(X, 1, K - 2)
    Line Input #4, xIn
    K = 0
    X = CSV_Scan(xIn, K)
    K = InStr(X, "File-")
    If K > 0 Then
        K1 = K + 5
        K2 = Len(X) - K + 4
        For K = 1 To Nb
            If X4 = arrDoc_Nom(K) Then
                blnOk = True
                NbOk = NbOk + 1
                If arrDoc_Id(K) = "" Then
                    arrDoc_Id(K) = Mid$(X, K1, K2)
                Else
                    Call MsgBox(X4 & " : " & arrDoc_Id(K), vbError, "Document déjà EXISTANT")
                End If
                Exit For
            End If
        Next K
       If Not blnOk Then Call MsgBox(X4, vbError, "Document inconnu")
    End If
    DoEvents
Loop
Call lstErr_AddItem(lstErr, cmdContext, "cmdJPL_SAB_DS : C:\Temp\SAB_Index.htm"): DoEvents
    
X = "<TABLE border = 1  width=1200 height=5 bgcolor=#0000FF cellpadding=3 ><TR>" _
         & "<TD  width=50 height=5><span style='font-size:10.0pt;font-family:Arial'><Font color=#FFFFFF>Code</TD>" _
         & "<TD  width=300 height=5><span style='font-size:10.0pt;font-family:Arial'><Font color=#FFFFFF>Produit</TD>" _
         & "<TD  width=650 height=5><span style='font-size:10.0pt;font-family:Arial'><Font color=#FFFFFF>Titre</TD>" _
         & "<TD  width=200 height=5><span style='font-size:10.0pt;font-family:Arial'><Font color=#FFFFFF>Document</TD>" _
        & "</TR></TABLE>"
Print #wIdFile_Destination, X
For I = 1 To Nb
    If I Mod 2 = 0 Then
        X = "<TABLE  width=1200 height=5 bgcolor=#F8F8FF cellpadding=3 ><TR>"
    Else
        X = "<TABLE  width=1200 height=5 bgcolor=#FFFFFF cellpadding=3 ><TR>"
    End If
    X1 = Mid$(arrProduit(I), 1, 3)
    X2 = Mid$(arrProduit(I), 4, Len(arrProduit(I)) - 3)
    X4 = "<A href=" & Asc34 & "http://docsrv:8080/docushare/dsweb/Get/Document-" & arrDoc_Id(I) & "/" & arrDoc_Nom(I) & Asc34 & ">" & arrDoc_Nom(I) & "</A>"
     X = X _
         & "<TD  width=50 height=5><span style='font-size:10.0pt;font-family:Arial'><Font color=#6060FF>" & X1 & "</span/TD>" _
         & "<TD  width=300 height=5><span style='font-size:10.0pt;font-family:Arial'><Font color=#6060FF>" & X2 & "</span/TD>" _
         & "<TD  width=650 height=5><span style='font-size:10.0pt;font-family:Arial'><Font color=#0000FF>" & arrDoc_Titre(I) & "</span/TD>" _
         & "<TD  width=200 height=5><span style='font-size:10.0pt;font-family:Arial'><Font color=#6060FF>" & X4 & "</TD>" _
        & "</TR></TABLE>"
   Print #wIdFile_Destination, X
    
   ' If NbOk Mod 1000 Then Call lstErr_ChangeLastItem(lstErr, cmdPrint, "cmdJPL_Import_INSEE: " & NbOk & "/" & Nb): DoEvents
    DoEvents
Next I


Close
Call lstErr_AddItem(lstErr, cmdPrint, "cmdJPL_SAB_DS " & NbOk & "/" & Nb): DoEvents

Exit Sub
'---------------------------------------------------------
Error_Handler:
'---------------------------------------------------------

Close

End Sub

Public Sub cmdJPL_BICBIC()
Dim X As String, K As Integer, Nb As Long, NbOk As Long
Dim xIn As String
Dim mAmj7 As Long
Dim blnSelect As Boolean
Dim xZTCHCOR0 As typeZTCHCOR0
Dim V
Dim xSQL As String
Dim xWhere As String, xSet As String, xValues As String
Dim X1 As String, X2 As String, X3 As String, X4 As String, X5 As String
Dim kErr As Integer
Dim blnDevise As Boolean
On Error GoTo Error_Handler

Nb = 0: NbOk = 0
rsZTCHCOR0_Init xZTCHCOR0
cnAdo.Open paramODBC_DSN_SAB
cnSab_Update.Open paramODBC_DSN_SAB
Call FEU_ROUGE

Open Trim("C:\Temp\SAB\wTCHCOR0.csv") For Output As #4
xSQL = "select * from BIAFIL.BICIBANP0 order by DEVISE"
Set rsSab = cnsab.Execute(xSQL)

    
Do Until rsSab.EOF
    Nb = Nb + 1
    X1 = Trim(rsSab("BICID"))
    X2 = rsSab("DEVISE")
    X3 = Trim(rsSab("IBANBICID"))
    X4 = ""
    X5 = ""
    If Len(X1) = 8 Then X1 = X1 & "XXX"
    If Len(X3) = 8 Then X3 = X3 & "XXX"
    kErr = 0
    Select Case X2
        Case "978": X2 = "EUR"
        Case "400": X2 = "USD"
        Case "006": X2 = "GBP"
        Case "008": X2 = "DKK"
        Case "028": X2 = "NOK"
        Case "036": X2 = "CHF"
        Case "404": X2 = "CAD"
        Case "732": X2 = "JPY"
        Case Else: blnDevise = False: kErr = 1
    End Select
    If kErr = 0 Then
        xSQL = "select SWIBICIN1 from " & paramIBM_Library_SAB & ".ZSWIBIC0 where SWIBICBIC = '" & X1 & "'"
        Set rsado = cnAdo.Execute(xSQL)
        If rsado.EOF Then
            kErr = 2:
        Else
            X4 = rsado("SWIBICIN1")
            xSQL = "select SWIBICIN1 from " & paramIBM_Library_SAB & ".ZSWIBIC0 where SWIBICBIC = '" & X3 & "'"
            Set rsado = cnAdo.Execute(xSQL)
            If rsado.EOF Then
                kErr = 3:
            Else
                X5 = rsado("SWIBICIN1")
                NbOk = NbOk + 1
                xZTCHCOR0.TCHCORBIC = X1
                xZTCHCOR0.TCHCORDEV = X2
                xZTCHCOR0.TCHCORBI1 = X3
                V = sqlZTCHCOR0_Insert(xZTCHCOR0)
                'à faire maj
            End If
        End If
    End If
    
    Print #4, kErr & ";" & X1 & ";" & X2 & ";" & X3 & ";" & X4 & ";" & X5
    rsSab.MoveNext

Loop



cnAdo.Close
Close
Call FEU_VERT
Call lstErr_AddItem(lstErr, cmdPrint, "cmdJPL_Import_DCN: " & NbOk & "/" & Nb): DoEvents

Exit Sub
'---------------------------------------------------------
Error_Handler:
'---------------------------------------------------------

Close
Call MsgBox(Err & " : " & Error(Err), vbCritical, "cmdJPL_Import_DCN")


End Sub
