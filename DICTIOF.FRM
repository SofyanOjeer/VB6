VERSION 5.00
Begin VB.Form frmDictio 
   Caption         =   "Gestion AS400 / DictioP1"
   ClientHeight    =   4905
   ClientLeft      =   225
   ClientTop       =   1365
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4905
   ScaleWidth      =   6615
   Begin VB.Frame fraCase 
      Caption         =   "Conversion Min / Maj"
      Height          =   975
      Left            =   240
      TabIndex        =   11
      Top             =   1080
      Width           =   6135
      Begin VB.OptionButton optCaseLU 
         Caption         =   "StrConv"
         Height          =   255
         Left            =   2160
         TabIndex        =   15
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton optCaseU 
         Caption         =   "majuscule"
         Height          =   375
         Left            =   2160
         TabIndex        =   14
         Top             =   120
         Width           =   1455
      End
      Begin VB.OptionButton optCaseL 
         Caption         =   "minuscule"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   600
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton optCaseNo 
         Caption         =   "néant"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmd919 
      Caption         =   "cmd 19 => 919"
      Height          =   375
      Left            =   2520
      TabIndex        =   10
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton cmdDevCrs 
      BackColor       =   &H008080FF&
      Caption         =   "Init DevCrs"
      Height          =   375
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   6200
      Picture         =   "DICTIOF.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   0
      Width           =   400
   End
   Begin VB.CommandButton cmdSupMdb 
      Caption         =   "Sup Mdb"
      Height          =   300
      Left            =   3000
      TabIndex        =   7
      Top             =   0
      Width           =   1000
   End
   Begin VB.CommandButton cmdReadMdb 
      Caption         =   "Lire Mdb"
      Default         =   -1  'True
      Height          =   300
      Left            =   2000
      TabIndex        =   6
      Top             =   0
      Width           =   1000
   End
   Begin VB.CommandButton cmdMàjMdb 
      Caption         =   "màj  Mdb"
      Height          =   300
      Left            =   1000
      TabIndex        =   5
      Top             =   0
      Width           =   1000
   End
   Begin VB.CommandButton cmdRechercher 
      Caption         =   "AS400 "
      Height          =   300
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   1000
   End
   Begin VB.ListBox LstDictio 
      Height          =   2010
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   6270
   End
   Begin VB.TextBox txtDicRub 
      Height          =   300
      Left            =   1080
      TabIndex        =   2
      Top             =   540
      Width           =   735
   End
   Begin VB.Label lblDicRub 
      Caption         =   "Rubrique"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   105
      TabIndex        =   1
      Top             =   585
      Width           =   855
   End
   Begin VB.Label lblErrMsg 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ErrMsg"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   4560
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "frmDictio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim arrTag() As Boolean, arrTagNb As Integer, lstErrClear As Boolean

Private recdictio As typeDictio
Private DictioMethod As String * 12
Private mDicrub As Integer
'---------------------------------------------------------
Public Sub Msg_Rcv(X As String)
'---------------------------------------------------------


End Sub


'---------------------------------------------------------
Private Sub lstDictioAddItem()
'---------------------------------------------------------
Dim I As Integer, K As Integer

For K = LstDictio.ListCount + 1 To arrDictioNb
LstDictio.AddItem arrDictio(K).DicCode & " " & arrDictio(K).DicLib _
             & Chr$(9) & arrDictio(K).DicTxt
Next K

LstDictio.ListIndex = LstDictio.ListCount - 1

End Sub


Private Sub chkCvMajMin_Click()

End Sub

Private Sub cmd919_Click()
Dim I As Integer, X As Variant

txtDicRub = 919
cmdSupMdb_Click

txtDicRub = 19
cmdReadMdb_Click

For I = 0 To LstDictio.ListCount - 1
    LstDictio.ListIndex = I
    recdictio.Method = "Seek="
    recdictio.DicCode = mId$(LstDictio.Text, 1, 4)
    recdictio.DicRub = "19"
    X = dbDictioRead(recdictio)
    If recdictio.Err = 0 Then
        recdictio.Method = constAddNew
        Mid$(recdictio.DicTxt, 231, 4) = mId$(recdictio.DicCode, 1, 4)
        recdictio.DicCode = mId$(recdictio.DicTxt, 7, 2)
        recdictio.DicRub = "919"
        If Trim(recdictio.DicCode) <> "" Then Call dbDictioUpdate(recdictio)

    End If
Next I

End Sub

Private Sub cmdDevCrs_Click()
'''' dbDeviseChange_Replication '"00000000"
End Sub

Private Sub cmdMàjMdb_Click()
Dim K As Integer
For K = 1 To arrDictioNb
    arrDictio(K).Method = constAddNew
    arrDictio(K).DicCode = Trim(arrDictio(K).DicCode)
    If optCaseL Then arrDictio(K).DicLib = LCase(arrDictio(K).DicLib)
    If optCaseU Then arrDictio(K).DicLib = UCase(arrDictio(K).DicLib)
    If optCaseLU Then arrDictio(K).DicLib = StrConv(arrDictio(K).DicLib, vbProperCase)
    Call dbDictioUpdate(arrDictio(K))
Next K
lstDictioEffacer

End Sub

'---------------------------------------------------------
Private Sub cmdQuit_Click()
'---------------------------------------------------------
Unload Me
End Sub

Private Sub cmdPrint_Click()
Dim Msg As String

Msg = Trim(txtDicRub)
prtDictioX Msg

End Sub

'---------------------------------------------------------
Private Sub cmdReadMdb_Click()
'---------------------------------------------------------
Dim X

lstDictioEffacer

recdictio.Method = "Seek>=      "
recdictio.DicCode = String$(20, Chr$(0))
recdictio.DicRub = Val(txtDicRub)
mDicrub = recdictio.DicRub

X = dbDictioRead(recdictio)
recdictio.Method = "MoveNext    "
Do While recdictio.Err = 0 _
     And recdictio.DicRub = mDicrub

    LstDictio.AddItem recdictio.DicCode & " " & recdictio.DicLib _
        & Chr$(9) & recdictio.DicTxt

X = dbDictioRead(recdictio)

Loop

End Sub

'---------------------------------------------------------
Private Sub XXX()
'---------------------------------------------------------
Dim X
recdictio.Method = "MoveFirst"

X = dbDictioRead(recdictio)
Do While recdictio.Err = 0 _

recdictio.Method = constUpdate
recdictio.DicLib = StrConv(recdictio.DicLib, vbProperCase)
X = dbDictioUpdate(recdictio)

recdictio.Method = "MoveNext    "
X = dbDictioRead(recdictio)

Loop


End Sub


'---------------------------------------------------------
Private Sub cmdRechercher_Click()
'---------------------------------------------------------

DictioMethod = "SnapP1      "
lstDictioEffacer
lstDictioRechercher

Do Until Not arrDictioSuite
    lstDictioSuite
Loop
End Sub

'---------------------------------------------------------
Public Sub lstDictioEffacer()
'---------------------------------------------------------
LstDictio.Clear
arrDictioNb = 0
recDictioInit recdictio
arrDictioSuite = True
End Sub


Private Sub cmdSupMdb_Click()
'---------------------------------------------------------
Dim X

lstDictioEffacer

recdictio.Method = "Seek>=      "
recdictio.DicCode = String$(20, Chr$(0))
recdictio.DicRub = Val(txtDicRub)
mDicrub = recdictio.DicRub

X = dbDictioRead(recdictio)
Do While recdictio.Err = 0 _
     And recdictio.DicRub = mDicrub

    recdictio.Method = constDelete
    X = dbDictioUpdate(recdictio)


    recdictio.Method = "MoveNext    "
    X = dbDictioRead(recdictio)

Loop


End Sub

'---------------------------------------------------------
Private Sub Form_Activate()
'---------------------------------------------------------
Set XForm = Me

txtDicRub.SetFocus

End Sub

'---------------------------------------------------------
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'---------------------------------------------------------
Select Case KeyCode
    Case Is = 27: cmdQuit_Click
    Case Is = 44: KeyCode = 0: frmElpPrt.prtScreen
    Case Is = 13: SendKeys "{TAB}"
End Select
lblErrMsg.Visible = False
End Sub

'---------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------

Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)

cmdDevCrs.Enabled = False

MDB_Master
DictioMethod = ""
Call lstDictioEffacer

End Sub


'---------------------------------------------------------
Private Sub Form_Unload(Cancel As Integer)
'---------------------------------------------------------
On Error Resume Next
MDB_Local

frmElp.SetFocus

End Sub


Private Sub txtDicRub_GotFocus()
'-------------------------------------------------'
Call txtGetFocus(txtDicRub)


End Sub


Private Sub txtDicRub_KeyPress(KeyAscii As Integer)
'-------------------------------------------------'
Call elpNum(txtDicRub, KeyAscii, 3, 0)


End Sub


Private Sub txtDicRub_LostFocus()
'-------------------------------------------------'
txtDicRub.BackColor = txtUsr.BackColor

End Sub


'---------------------------------------------------------
Public Sub lstDictioRechercher()
'---------------------------------------------------------

lstDictioSql

If IsNull(srvDictioMon(recdictio)) Then
    If mId$(DictioMethod, 1, 4) = "Seek" Then
        Call arrDictioAddItem(recdictio)
        Call lstDictioAddItem
    Else
        Call lstDictioAddItem
    End If
End If

End Sub


'---------------------------------------------------------
Private Sub lstDictioSuite()
'---------------------------------------------------------
Dim X As String

If arrDictioNb > 0 Then
    recdictio = arrDictio(arrDictioNb)
    X = "+"
Else
    lstDictioSql
    X = " "
End If
    
If Trim(mId$(DictioMethod, 5, 2)) = "" Then Mid$(DictioMethod, 5, 2) = "L1"

recdictio.Method = "Snap" & mId$(DictioMethod, 5, 2) & X
 
If IsNull(srvDictioMon(recdictio)) Then
    Call lstDictioAddItem
End If

End Sub

'---------------------------------------------------------
Private Sub lstDictioSql()
'---------------------------------------------------------

recdictio.Method = DictioMethod
recdictio.DicRub = Format$(Val(txtDicRub), "000")
recdictio.DicCode = Space$(20)

arrDictio(0) = recdictio
arrDictio(0).DicCode = "99999999999"
End Sub






