VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmKHALDZAL 
   AutoRedraw      =   -1  'True
   Caption         =   "KHALDZAL"
   ClientHeight    =   10530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13575
   Icon            =   "KHALDZAL_F.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   10530
   ScaleWidth      =   13575
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6120
      TabIndex        =   4
      Top             =   0
      Width           =   6900
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9852
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   13500
      _ExtentX        =   23813
      _ExtentY        =   17383
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "@Arial Unicode MS"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "KHALDZAL"
      TabPicture(0)   =   "KHALDZAL_F.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraTab0"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Paramétrage"
      TabPicture(1)   =   "KHALDZAL_F.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstW"
      Tab(1).ControlCount=   1
      Begin VB.ListBox lstW 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1860
         Left            =   -67320
         TabIndex        =   9
         Top             =   6360
         Visible         =   0   'False
         Width           =   4212
      End
      Begin VB.Frame fraTab0 
         Height          =   9420
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   13296
         Begin VB.Frame fraSelect_Options 
            BackColor       =   &H00F0FFFF&
            Height          =   555
            Left            =   360
            TabIndex        =   14
            Top             =   240
            Visible         =   0   'False
            Width           =   8712
         End
         Begin VB.Frame fraDetail 
            BackColor       =   &H00E0FFFF&
            Height          =   3390
            Left            =   4095
            TabIndex        =   10
            Top             =   1140
            Visible         =   0   'False
            Width           =   7932
            Begin VB.Label libSAAMSGMTD 
               BackColor       =   &H00E0FFFF&
               Caption         =   "SAAMSGDTRT"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3405
               TabIndex        =   15
               Top             =   120
               Width           =   2325
            End
            Begin VB.Label libSAAMSGDTRT 
               BackColor       =   &H00E0FFFF&
               Caption         =   "SAAMSGDTRT"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1290
               TabIndex        =   13
               Top             =   105
               Width           =   1755
            End
            Begin VB.Label libSAAMSGTYPE 
               BackColor       =   &H00E0FFFF&
               Caption         =   "SAAMSGTYPE"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   105
               TabIndex        =   12
               Top             =   105
               Width           =   975
            End
            Begin VB.Label libSAAMSGTXT 
               BackColor       =   &H00FFFFFF&
               Caption         =   "SAAMSGTXT"
               BeginProperty Font 
                  Name            =   "@Arial Unicode MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2805
               Left            =   120
               TabIndex        =   11
               Top             =   465
               Width           =   7575
               WordWrap        =   -1  'True
            End
         End
         Begin VB.ComboBox cboSelect_SQL 
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   324
            Left            =   9240
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   240
            Width           =   3732
         End
         Begin VB.CommandButton cmdSelect_Ok 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Rechercher"
            Height          =   555
            Left            =   10440
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   720
            Width           =   1335
         End
         Begin MSFlexGridLib.MSFlexGrid fgSelect 
            Height          =   7668
            Left            =   360
            TabIndex        =   5
            Top             =   1560
            Visible         =   0   'False
            Width           =   3312
            _ExtentX        =   5847
            _ExtentY        =   13520
            _Version        =   393216
            Rows            =   1
            Cols            =   7
            FixedCols       =   0
            RowHeightMin    =   350
            BackColor       =   -2147483633
            ForeColor       =   12582912
            BackColorFixed  =   8421376
            ForeColorFixed  =   -2147483637
            BackColorSel    =   12648384
            BackColorBkg    =   -2147483633
            AllowBigSelection=   0   'False
            FocusRect       =   2
            HighLight       =   0
            GridLinesFixed  =   1
            AllowUserResizing=   3
            FormatString    =   "> SAAMsgId             | HISMVTID                ||"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Unicode MS"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSFlexGridLib.MSFlexGrid fgDetail 
            Height          =   3675
            Left            =   3660
            TabIndex        =   8
            Top             =   5175
            Visible         =   0   'False
            Width           =   9375
            _ExtentX        =   16536
            _ExtentY        =   6482
            _Version        =   393216
            Cols            =   9
            FixedCols       =   0
            BackColor       =   16777215
            BackColorFixed  =   8421376
            ForeColorFixed  =   -2147483633
            BackColorBkg    =   -2147483633
            FormatString    =   $"KHALDZAL_F.frx":0342
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "@Arial Unicode MS"
               Size            =   7.5
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
      Picture         =   "KHALDZAL_F.frx":03F2
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
   Begin VB.Menu mnuPrint 
      Caption         =   "mnuPrint"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmKHALDZAL"
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
Dim YSAAMSG0_Aut As typeAuthorization
Dim blnAuto As Boolean, blnError As Boolean
Dim cmdSelect_SQL_K As String

Dim fgSelect_FormatString As String, fgSelect_K As Integer
Dim fgSelect_RowDisplay As Integer, fgSelect_RowClick As Integer, fgSelect_ColClick As Integer
Dim fgSelect_ColorClick As Long, fgSelect_ColorDisplay As Long
Dim fgSelect_Sort1 As Integer, fgSelect_Sort2 As Integer
Dim fgSelect_SortAD As Integer, fgSelect_Sort1_Old As Integer
Dim fgSelect_arrIndex As Integer
Dim blnfgSelect_DisplayLine As Boolean


'______________________________________________________________________

Dim wAmjMin As String, wAmjMax As String, wHmsMin As Long, wHmsMax As Long
Dim xYSAAMSG0 As typeYSAAMSG0, newYSAAMSG0 As typeYSAAMSG0, oldYSAAMSG0 As typeYSAAMSG0
Dim arrYSAAMSG0() As typeYSAAMSG0, arrYSAAMSG0_Nb As Long, arrYSAAMSG0_Max As Long, arrYSAAMSG0_Index As Long

Dim xKHALDZAL As typeKHALDZAL, newKHALDZAL As typeKHALDZAL, oldKHALDZAL As typeKHALDZAL
Dim arrKHALDZAL() As typeKHALDZAL, arrKHALDZAL_Nb As Long, arrKHALDZAL_Max As Long, arrKHALDZAL_Index As Long

Dim fgDetail_FormatString As String, fgDetail_K As Integer
Dim fgDetail_RowDisplay As Integer, fgDetail_RowClick As Integer, fgDetail_ColClick As Integer
Dim fgDetail_ColorClick As Long, fgDetail_ColorDisplay As Long
Dim fgDetail_Sort1 As Integer, fgDetail_Sort2 As Integer
Dim fgDetail_SortAD As Integer, fgDetail_Sort1_Old As Integer
Dim fgDetail_arrIndex As Integer
Dim blnfgDetail_DisplayLine As Boolean
Dim xFF As String

Dim xYSAAMVTLNK As typeYSAAMVTLNK, newYSAAMVTLNK As typeYSAAMVTLNK, oldYSAAMVTLNK As typeYSAAMVTLNK
Dim xYSAAMSG1 As typeYSAAMSG1, newYSAAMSG1 As typeYSAAMSG1, oldYSAAMSG1 As typeYSAAMSG1

'______________________________________________________________________
Private Sub fgSelect_Display()
Dim wColor As Long
Dim xSql As String
Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
SSTab1.Tab = 0
fgSelect.Visible = False
fgSelect_Reset

fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
fgSelect.Row = 0

currentAction = "fgSelect_Display"
KHALDZAL.cnSAB073Y_Open

xSql = "select * from YSAAMVTLNK  " _
     & " order by SAAMSGID"
Set rsYSAAMVTLNK = cnSAB073Y.Execute(xSql)

Do While Not rsYSAAMVTLNK.EOF
     V = rsYSAAMVTLNK_GetBuffer(rsYSAAMVTLNK, xYSAAMVTLNK)

    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1
    fgSelect_DisplayLine I, True
    
    rsYSAAMVTLNK.MoveNext
Loop

fgSelect.Visible = True

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & arrYSAAMSG0_Nb): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub
Private Sub arrYSAAMSG0_SQL(xWhere As String)
Dim V
Dim X As String, xSql As String, K As Integer
On Error GoTo Error_Handler
ReDim arrYSAAMSG0(101)
arrYSAAMSG0_Max = 100: arrYSAAMSG0_Nb = 0

Set rsSab = Nothing

xSql = "select * from " & paramIBM_Library_SABSPE & ".YSAAMSG0 " & xWhere
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    V = rsYSAAMSG0_GetBuffer(rsSab, xYSAAMSG0)
    
     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "frmYSAAMSG0.fgselect_Display"
        '' Exit Sub
     Else
         arrYSAAMSG0_Nb = arrYSAAMSG0_Nb + 1
         If arrYSAAMSG0_Nb > arrYSAAMSG0_Max Then
             arrYSAAMSG0_Max = arrYSAAMSG0_Max + 100
             ReDim Preserve arrYSAAMSG0(arrYSAAMSG0_Max)
         End If
         
         arrYSAAMSG0(arrYSAAMSG0_Nb) = xYSAAMSG0
    End If
    rsSab.MoveNext
Loop


Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub


Public Sub cmdSelect_Reset()
If blnControl Then
    lstErr.Clear
    fgSelect.Visible = False
    fgDetail.Visible = False: fraDetail.Visible = False
    lstW.Visible = False
    cmdSelect_Ok.Visible = True
    cmdSelect_SQL_K = Trim(Mid$(cboSelect_SQL, 1, 3))

End If

End Sub



Private Sub fgDetail_Display()
Dim wColor As Long
Dim xWhere As String, X As String

Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean
Dim xSql As String

On Error GoTo Error_Handler
fgDetail.Visible = False: fraDetail.Visible = False
fgDetail_Reset

fgDetail.Rows = 1
fgDetail.FormatString = fgDetail_FormatString
fgDetail.Row = 0

currentAction = "fgDetail_Display"

xSql = "select * from YSAAMSG0  " _
     & " where  SAAMSGID = " & oldYSAAMVTLNK.SAAMSGID
Set rsYSAAMSG0 = cnSAB073Y.Execute(xSql)

If Not rsYSAAMSG0.EOF Then
    libSAAMSGDTRT = dateImp10(rsYSAAMSG0("SAAMsgDTrt"))
    libSAAMSGTYPE = rsYSAAMSG0("SAAMsgType")
    oldYSAAMSG0.SAAMsgDev = rsYSAAMSG0("SAAMsgDev")
    oldYSAAMSG0.SAAMsgMt = -(rsYSAAMSG0("SAAMsgMt"))
    libSAAMSGMTD = Format$(oldYSAAMSG0.SAAMsgMt, "### ### ### ##0.00-") & " " & oldYSAAMSG0.SAAMsgDev
End If

xSql = "select * from YSAAMSG1  where SAAMsgId = " & oldYSAAMVTLNK.SAAMSGID _
     & " order by SAAMsgSeq"
Set rsYSAAMSG1 = cnSAB073Y.Execute(xSql)
X = ""
Do While Not rsYSAAMSG1.EOF
    V = rsYSAAMSG1_GetBuffer(rsYSAAMSG1, xYSAAMSG1)
    X = X & xYSAAMSG1.SAAMsgFld & xYSAAMSG1.SAAMsgFldX & ": " & xYSAAMSG1.SAAMSGTXT & vbCrLf
        rsYSAAMSG1.MoveNext
Loop
libSAAMSGTXT = X



xSql = "select * from KHALDZAL  where SAAMsgId = 0 and HISMVTMTD > " & cur_19P(oldYSAAMSG0.SAAMsgMt - 0.5) _
     & " and HISMVTMTD < " & cur_19P(oldYSAAMSG0.SAAMsgMt + 0.5) & " and HISMVTDEV = '" & oldYSAAMSG0.SAAMsgDev & "' order by id"
Set rsKHALDZAL = cnSAB073Y.Execute(xSql)
X = ""
Do While Not rsKHALDZAL.EOF
     V = rsKHALDZAL_GetBuffer(rsKHALDZAL, xKHALDZAL)

    fgDetail.Rows = fgDetail.Rows + 1
    fgDetail.Row = fgDetail.Rows - 1
    fgDetail_DisplayLine I
    
    rsKHALDZAL.MoveNext
Loop

fgDetail.Visible = True: fraDetail.Visible = True


Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Public Sub fgSelect_DisplayLine(lIndex As Long, blnYSAAMSG0 As Boolean)
Dim K As Integer
Dim wColor As Long, wColor_Row As Long
Dim xSql As String
On Error Resume Next

 Select Case xYSAAMVTLNK.HISMVTID
    Case Is > 0: wColor = RGB(128, 128, 128)
    Case Else: wColor = RGB(64, 64, 128)
End Select

fgSelect.Col = 0: fgSelect.Text = xYSAAMVTLNK.SAAMSGID
fgSelect.CellForeColor = wColor
fgSelect.Col = 1: fgSelect.Text = xYSAAMVTLNK.HISMVTID
fgSelect.CellForeColor = wColor


fgSelect.Col = fgSelect_arrIndex: fgSelect.Text = lIndex
End Sub
Public Sub fgDetail_DisplayLine(lIndex As Long)
Dim K As Integer
Dim wColor As Long, wColor_Row As Long

On Error Resume Next
wColor = vbBlue: wColor_Row = vbWhite
fgDetail.Col = 0: fgDetail.Text = xKHALDZAL.Id
fgDetail.Col = 1: fgDetail.Text = dateImp10(xKHALDZAL.HISMVTDTRT)
fgDetail.Col = 2: fgDetail.Text = xKHALDZAL.HISMVTOPEC
fgDetail.Col = 3: fgDetail.Text = Format$(xKHALDZAL.HISMVTMTD, "### ### ### ##0.00")
fgDetail.Col = 4: fgDetail.Text = xKHALDZAL.HISMVTLIB1

fgDetail.Col = fgDetail_arrIndex: fgDetail.Text = lIndex
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
Dim I As Integer, X As String, wIndex As Long

For I = 1 To fgSelect.Rows - 1
    fgSelect.Row = I
    fgSelect.Col = fgSelect_arrIndex
    wIndex = Val(fgSelect.Text)
    Select Case lK
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


Public Sub Msg_Rcv(Msg As String)
'---------------------------------------------------------
Dim wFct As String
mWindowState = Me.WindowState
If Me.WindowState <> frmUsr_Windowstate Then Me.WindowState = frmUsr_Windowstate

wFct = UCase$(Trim(Mid$(Msg, 1, 12)))
Call BiaPgmAut_Init(wFct, YSAAMSG0_Aut)

'blnSetfocus = True
Form_Init


Select Case wFct
    Case Else: blnAuto = False
End Select

End Sub


Public Sub Form_Init()
Dim V, xSql As String, X As String
Dim K As Long

Me.Enabled = False
Call lstErr_Clear(lstErr, cmdPrint, "Initialisation ")
DoEvents
lstErr.Visible = True

blnControl = False

cmdReset

libSAAMSGMTD.ForeColor = vbMagenta
'libSAAMSGTXT.FontSize = 7
xFF = Chr$(159) & Chr$(159)
fgSelect_FormatString = fgSelect.FormatString
fgSelect.Enabled = True
fgSelect.Visible = False




lstW.Visible = False

fgDetail.Visible = False: fraDetail.Visible = False
fgDetail_FormatString = fgDetail.FormatString

cboSelect_SQL.Clear
    cboSelect_SQL.AddItem "K0 - YSAAMSG init + import"
    cboSelect_SQL.AddItem "K0b - YSAAMSG - 201 => 2*1"
    cboSelect_SQL.AddItem "K1 - HISMVTP0 import = > KHALDZAL"
    cboSelect_SQL.AddItem "K2 - match => KHALDZAL"
    cboSelect_SQL.AddItem "K2b - match NOK =>(YSAAMVTLNK)"
    cboSelect_SQL.AddItem "K2c - Rapprochement manuel"
    cboSelect_SQL.AddItem "K2d - YSAAMVTLNK => KHALDZAL"
    cboSelect_SQL.AddItem "K3 - YSAAMSG1 => KHALDZAL"
    cboSelect_SQL.AddItem "K3* - 201 : YSAAMSG1 => KHALDZAL"
    cboSelect_SQL.AddItem "K4i - PDF_X => KHALDZAL"
    cboSelect_SQL.AddItem "K4r - PDF (reprise) => KHALDZAL"
cboSelect_SQL.ListIndex = 0

lstW.Clear


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
blnControl = True

End Sub



Public Sub fgSelect_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
fgSelect.Visible = False
mRow = fgSelect.Row

If lRow > 0 And lRow < fgSelect.Rows Then
    fgSelect.Row = lRow
    For I = fgSelect_arrIndex To fgSelect.FixedCols Step -1
        fgSelect.Col = I: fgSelect.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgSelect.Row = mRow
    If fgSelect.Row > 0 Then
        lRow = fgSelect.Row
        lColor_Old = fgSelect.CellBackColor
        For I = fgSelect_arrIndex To fgSelect.FixedCols Step -1
          fgSelect.Col = I: fgSelect.CellBackColor = lColor
        Next I
    End If
End If
fgSelect.LeftCol = fgSelect.FixedCols
fgSelect.Visible = True
End Sub

Public Sub fgDetail_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
fgDetail.Visible = False: fraDetail.Visible = False
mRow = fgDetail.Row

If lRow > 0 And lRow < fgDetail.Rows Then
    fgDetail.Row = lRow
    For I = fgDetail_arrIndex To fgDetail.FixedCols Step -1
        fgDetail.Col = I: fgDetail.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgDetail.Row = mRow
    If fgDetail.Row > 0 Then
        lRow = fgDetail.Row
        lColor_Old = fgDetail.CellBackColor
        For I = fgDetail_arrIndex To fgDetail.FixedCols Step -1
          fgDetail.Col = I: fgDetail.CellBackColor = lColor
        Next I
    End If
End If
fgDetail.LeftCol = fgDetail.FixedCols
fgDetail.Visible = True: fraDetail.Visible = True
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







Private Sub cboSelect_SQL_Click()
cmdSelect_Reset

End Sub


Private Sub cmdContext_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Select Case cmdContext.Caption
    Case Is = constcmdRechercher: Me.PopupMenu mnuContext, vbPopupMenuLeftButton
    Case Is = constcmdAbandonner: cmdContext_Quit
End Select

End Sub

Private Sub cmdPrint_Click()
Dim X As String, I As Integer
Me.Enabled = False: Me.MousePointer = vbHourglass

Select Case SSTab1.Tab
    Case 0:
        Me.PopupMenu mnuPrint, vbPopupMenuLeftButton
    End Select

Me.Enabled = True: Me.MousePointer = 0
End Sub

Private Sub cmdSelect_Ok_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> KHAMDZAL_cmdSelect_Ok ........"): DoEvents

If fgSelect.Visible Then cmdSelect_Reset
fgSelect.Visible = False
fraSelect_Options.Visible = False

Select Case cmdSelect_SQL_K
    Case "K0":  KHALDZAL.YSAAMSG_Init
    Case "K0b":  KHALDZAL.YSAAMSG_201
    Case "K1":  KHALDZAL.HISMVTP0_Import
    Case "K2":  KHALDZAL.HISMVTP0_Match
    Case "K2b":  KHALDZAL.HISMVTP0_match_NOK
    Case "K2c":  fgSelect_Display
    Case "K2d":  YSAAMVTLNK_KHALDZAL
    Case "K3":  KHALDZAL.HISMVTP0_YSAAMSG1
    Case "K3*": YSAAMSG_201_Origine
    Case "K4i":  ScanLink_Init
    Case "K4r":  ScanLink_Reprise
End Select
    
    
Call lstErr_AddItem(lstErr, cmdContext, "< KHAMDZAL_cmdSelect_Ok"): DoEvents
Me.Enabled = True: Me.MousePointer = 0
If cmdSelect_Ok.Visible Then cmdSelect_Ok.SetFocus

End Sub


Private Sub fgDetail_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim wOrigine As String

Dim xSql As String
On Error GoTo Error_Handle


If y <= fgDetail.RowHeightMin Then
Else
    If fgDetail.Rows > 1 Then
       ' blnControl = False
        Call fgDetail_Color(fgDetail_RowClick, MouseMoveUsr.BackColor, fgDetail_ColorClick)
        fgDetail.Col = 0:  oldYSAAMVTLNK.HISMVTID = CLng(fgDetail.Text)
        xSql = "update YSAAMVTLNK  set HISMVTID = " & oldYSAAMVTLNK.HISMVTID _
             & " where SAAMSGID = " & oldYSAAMVTLNK.SAAMSGID
        Call FEU_ROUGE
        Set rsYSAAMVTLNK = cnSAB073Y.Execute(xSql)
        Call FEU_VERT
        
        fgSelect.Col = 1: fgSelect.Text = oldYSAAMVTLNK.HISMVTID

   End If
End If
Exit Sub

Error_Handle:
 MsgBox Error, vbCritical, "erreur : fgDetail_MouseDown"

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
    Case Is = 27: cmdContext_Quit: KeyCode = 0
'   Case Is = 34: cmdPageNext_Click
'   Case Is = 33: cmdPagePrior_Click
    Case Is = 44: KeyCode = 0: frmElpPrt.prtScreen
End Select


End Sub

Public Sub cmdContext_Quit()
'blnControl = False
lstErr.Clear: lstErr.Height = 200

If SSTab1.Tab <> 0 Then
    SSTab1.Tab = 0
    Exit Sub
End If

If fgDetail.Visible Then
    fgDetail.Visible = False: fraDetail.Visible = False
    Exit Sub
End If
If fgSelect.Visible Then
    fgSelect.Visible = False
    Exit Sub
End If

If SSTab1.Tab = 0 Then
    Unload Me
End If

End Sub
Public Sub cmdContext_Return()
    If SSTab1.Tab = 0 Then
        If cmdSelect_SQL_K <> "J" And cmdSelect_SQL_K <> "J#" Then
            If Not fgSelect.Version Then cmdSelect_Ok_Click
        End If
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
Dim wOrigine As String, xSql As String
On Error Resume Next


If y <= fgSelect.RowHeightMin Then
    Select Case fgSelect.Col
        Case 0: fgSelect_Sort1 = 0: fgSelect_Sort2 = 2: fgSelect_Sort
        Case 1:  fgSelect_Sort1 = 1: fgSelect_Sort2 = 2: fgSelect_Sort
        Case fgSelect_arrIndex:  fgSelect_SortX fgSelect_arrIndex
    End Select
Else
    If fgSelect.Rows > 1 Then
        Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
        fgSelect.Col = 0: oldYSAAMVTLNK.SAAMSGID = CLng(fgSelect.Text)
        
        fgDetail_Display
        
   End If
End If

End Sub

Public Sub fgSelect_Reset()
fgSelect.Clear
fgSelect_Sort1 = 0: fgSelect_Sort2 = 0
fgSelect_Sort1_Old = -1
fgSelect_RowDisplay = 0: fgSelect_RowClick = 0
fgSelect_arrIndex = fgSelect.Cols - 1
blnfgSelect_DisplayLine = False
fgSelect_SortAD = 6
fgSelect.LeftCol = fgSelect.FixedCols

End Sub


Public Sub fgDetail_Reset()
fgDetail.Clear
fgDetail_Sort1 = 0: fgDetail_Sort2 = 0
fgDetail_Sort1_Old = -1
fgDetail_RowDisplay = 0: fgDetail_RowClick = 0
fgDetail_arrIndex = fgDetail.Cols - 1
blnfgDetail_DisplayLine = False
fgDetail_SortAD = 6
fgDetail.LeftCol = fgDetail.FixedCols

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






Public Sub fgSelect_ForeColor(lColor As Long)
For I = 0 To fgSelect_arrIndex
  fgSelect.Col = I: fgSelect.CellForeColor = lColor
Next I

End Sub

































