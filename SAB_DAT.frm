VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form frmSAB_DAT 
   AutoRedraw      =   -1  'True
   Caption         =   "SAB_DAT"
   ClientHeight    =   10305
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   13530
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SAB_DAT.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   10305
   ScaleWidth      =   13530
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   480
      Left            =   6120
      TabIndex        =   2
      Top             =   0
      Width           =   6900
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9720
      Left            =   0
      TabIndex        =   3
      Top             =   495
      Width           =   13440
      _ExtentX        =   23707
      _ExtentY        =   17145
      _Version        =   393216
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Rechercher"
      TabPicture(0)   =   "SAB_DAT.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraSelect"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "SAB_DAT.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtFg"
      Tab(1).Control(1)=   "txtRTF"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "SAB_DAT.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.TextBox txtFg 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2790
         Left            =   -69030
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   8
         Text            =   "SAB_DAT.frx":035E
         Top             =   1155
         Visible         =   0   'False
         Width           =   5775
      End
      Begin VB.Frame fraSelect 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   9630
         Left            =   -135
         TabIndex        =   4
         Top             =   495
         Width           =   13425
         Begin MSFlexGridLib.MSFlexGrid fgDetail 
            Height          =   7680
            Left            =   8805
            TabIndex        =   13
            Top             =   3570
            Visible         =   0   'False
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   13547
            _Version        =   393216
            FixedCols       =   0
            RowHeightMin    =   400
            BackColor       =   15790320
            ForeColor       =   4210752
            BackColorFixed  =   8421504
            ForeColorFixed  =   16777215
            BackColorBkg    =   15790320
            GridColor       =   10526720
            GridColorFixed  =   10526720
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            FormatString    =   "<Code           |<Intitulé                                                               "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.CommandButton cmdSelect_Ok 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Rechercher"
            Height          =   555
            Left            =   11820
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   705
            Width           =   1335
         End
         Begin VB.ComboBox cboSelect_SQL 
            Height          =   330
            Left            =   9840
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   300
            Width           =   3435
         End
         Begin VB.Frame fraSelect_Options 
            BackColor       =   &H00F0FFFF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1305
            Left            =   135
            TabIndex        =   5
            Top             =   120
            Visible         =   0   'False
            Width           =   9375
            Begin VB.ComboBox cboSelect_DATOPECLI 
               Height          =   330
               Left            =   1080
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   12
               Top             =   435
               Width           =   1860
            End
            Begin MSComCtl2.DTPicker txtSelect_DATOPEDIS 
               Height          =   300
               Left            =   5085
               TabIndex        =   14
               Top             =   450
               Visible         =   0   'False
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
               Format          =   89522179
               CurrentDate     =   36299
               MaxDate         =   401768
               MinDate         =   -328351
            End
            Begin VB.Label lblSelect_DATOPEDIS 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Date MAD >="
               Height          =   270
               Left            =   3765
               TabIndex        =   15
               Top             =   540
               Visible         =   0   'False
               Width           =   1080
            End
            Begin VB.Label lblSelect_DATOPECLI 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Client"
               Height          =   270
               Left            =   225
               TabIndex        =   11
               Top             =   480
               Width           =   750
            End
         End
         Begin MSFlexGridLib.MSFlexGrid fgSelect 
            Height          =   7710
            Left            =   90
            TabIndex        =   10
            Top             =   1380
            Width           =   13200
            _ExtentX        =   23283
            _ExtentY        =   13600
            _Version        =   393216
            Cols            =   15
            FixedCols       =   0
            RowHeightMin    =   400
            BackColor       =   16777215
            ForeColor       =   0
            BackColorFixed  =   8421376
            ForeColorFixed  =   16777215
            BackColorBkg    =   16777215
            GridColor       =   10526720
            GridColorFixed  =   10526720
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            FormatString    =   $"SAB_DAT.frx":0366
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin RichTextLib.RichTextBox txtRTF 
         Height          =   5610
         Left            =   -69525
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   3450
         Visible         =   0   'False
         Width           =   7755
         _ExtentX        =   13679
         _ExtentY        =   9895
         _Version        =   393217
         BackColor       =   15790320
         Enabled         =   -1  'True
         HideSelection   =   0   'False
         ScrollBars      =   3
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"SAB_DAT.frx":047C
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
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   13080
      Picture         =   "SAB_DAT.frx":04FC
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   500
   End
   Begin VB.Menu mnuContext 
      Caption         =   "mnuContext"
      Visible         =   0   'False
      Begin VB.Menu mnuContextQuitter 
         Caption         =   "Quitter"
      End
   End
   Begin VB.Menu mnuPrint 
      Caption         =   "mnuPrint"
      Visible         =   0   'False
      Begin VB.Menu mnuPrint_Mail 
         Caption         =   "Envoi Mail"
      End
      Begin VB.Menu mnuPrint_Excel 
         Caption         =   "Excel"
      End
   End
End
Attribute VB_Name = "frmSAB_DAT"
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
Dim arrHab(19) As Boolean
Dim blnAuto As Boolean, blnError As Boolean
Dim cmdSelect_SQL_K As String

Dim fgSelect_FormatString As String, fgSelect_K As Integer
Dim fgSelect_RowDisplay As Integer, fgSelect_RowClick As Integer, fgSelect_ColClick As Integer
Dim fgSelect_ColorClick As Long, fgSelect_ColorDisplay As Long
Dim fgSelect_Sort1 As Integer, fgSelect_Sort2 As Integer
Dim fgSelect_SortAD As Integer, fgSelect_Sort1_Old As Integer
Dim fgSelect_arrIndex As Integer
Dim blnfgSelect_DisplayLine As Boolean

Dim fgDetail_FormatString As String, fgDetail_K As Integer
Dim fgDetail_RowDisplay As Integer, fgDetail_RowClick As Integer, fgDetail_ColClick As Integer
Dim fgDetail_ColorClick As Long, fgDetail_ColorDisplay As Long
Dim fgDetail_Sort1 As Integer, fgDetail_Sort2 As Integer
Dim fgDetail_SortAD As Integer, fgDetail_Sort1_Old As Integer
Dim fgDetail_arrIndex As Integer
Dim blnfgDetail_DisplayLine As Boolean

'______________________________________________________________________

Dim wAMJMin As String, WAMJMax As String, wHmsMin As Long, wHmsMax As Long



Dim HeightOfLine As Long, LinesOfText As Long

Dim txtRTF_prtForeColor_Header As Long

Dim mDATOPECLI As String

Dim rsSabX As New ADODB.Recordset

Private Function retourne_encours_restant(zDossier As Long, zClient As String, zEtb As Long, ByRef zMontant As Double) As Boolean
Dim xSQL As String
Dim rs As ADODB.Recordset

    retourne_encours_restant = False
    zMontant = 0
    xSQL = "SELECT CAUAGARES FROM " & paramIBM_Library_SAB & ".ZCAUAGA0 WHERE CAUAGACLI = '" & zClient & "'"
    xSQL = xSQL & " AND CAUAGADOS = " & zDossier & " and CAUAGAETB = " & zEtb
    xSQL = xSQL & " ORDER BY CAUAGADAT DESC"
    Set rs = cnsab.Execute(xSQL)
    If Not rs.EOF Then
        zMontant = CDbl(rs("CAUAGARES"))
        retourne_encours_restant = True
    End If
    If rs.State = adStateOpen Then
        rs.Close
    End If
    Set rs = Nothing

End Function

Public Sub fgDetail_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
fgDetail.Visible = False
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
fgDetail.Visible = True
End Sub

Private Sub fgDetail_Display(lCIB As String)
Dim X As String, xWhere As String
Dim xSQL As String

On Error GoTo Error_Handler

currentAction = "fgDetail_Display"
fgDetail.Visible = False
fgDetail_Reset

fgDetail.Rows = 1
fgDetail.FormatString = "<Guichet  |<BIC                     |<Nom du guichet                      |<CP Ville                                              |<Adresse                                                                                                 |"
fgDetail.Row = 0
'___________________________________________________________________________

  
Do While Not rsSab.EOF

    fgDetail.Rows = fgDetail.Rows + 1
    fgDetail.Row = fgDetail.Rows - 1
    fgDetail_Display_Line
    
    rsSab.MoveNext

Loop

fgDetail.Visible = True

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgDetail.Rows - 1): DoEvents

'___________________________________________________________________________

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub



Public Sub fgDetail_Display_Line()
Dim X As String, wColor As Long

On Error Resume Next


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




Public Sub fgDetail_Sort()
If fgDetail.Rows > 1 Then
    fgDetail.Row = 1
    fgDetail.RowSel = fgDetail.Rows - 1
    
    If fgDetail_Sort1_Old = fgDetail_Sort1 Then
        If fgDetail_SortAD = 5 Then
            fgDetail_SortAD = 6
        Else
            fgDetail_SortAD = 5
        End If
    Else
        fgDetail_SortAD = 5
    End If
    fgDetail_Sort1_Old = fgDetail_Sort1
    
    fgDetail.Col = fgDetail_Sort1
    fgDetail.ColSel = fgDetail_Sort2
    fgDetail.Sort = fgDetail_SortAD
End If

End Sub



Public Sub Form_Init()
Dim V, xSQL As String, X As String
Dim K As Long

Me.Enabled = False
Call lstErr_Clear(lstErr, cmdPrint, "Initialisation ")
DoEvents
lstErr.Visible = True


cmdReset
blnControl = False

fgSelect_FormatString = fgSelect.FormatString
fgSelect.Enabled = True
fgSelect.Visible = False

fgDetail_FormatString = fgDetail.FormatString
fgDetail.Enabled = True
fgDetail.Visible = False
fgDetail.Top = fgSelect.Top
fgDetail.Left = 3500

cboSelect_DATOPECLI.Clear
cboSelect_DATOPECLI.AddItem " "
xSQL = "select distinct DATOPECLI from " & paramIBM_Library_SAB & ".ZDATOPE0  " _
     & " where DATOPEETA = '03'" _
     & " order by DATOPECLI"
Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF

    cboSelect_DATOPECLI.AddItem rsSab("DATOPECLI")
    rsSab.MoveNext

Loop

Call DTPicker_Set(txtSelect_DATOPEDIS, YBIATAB0_DATE_CPT_J)


fraSelect_Options.Visible = True

If cboSelect_SQL.ListCount > 0 Then cboSelect_SQL.ListIndex = 0

blnControl = True


cmdSelect_Reset
Me.Enabled = True

End Sub



'-------------------------------------------------------
Sub txt_GotFocus(C As Control)
'-------------------------------------------------------
currentActiveControl_Name = C.Name
C.ForeColor = txtUsr.ForeColor
C.BackColor = focusUsr.BackColor
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

'-------------------------------------------------------
Sub txt_LostFocus(C As Control)
'-------------------------------------------------------
arrTag(Val(C.Tag)) = True
C.ForeColor = txtUsr.ForeColor
C.BackColor = txtUsr.BackColor
End Sub


Public Sub fgSelect_SortX(lK As Integer)
Dim I As Integer, X As String, wIndex As Long

For I = 1 To fgSelect.Rows - 1
    fgSelect.Row = I
    fgSelect.Col = lK
    Select Case lK
'        Case 3: fgSelect.Col = 3: X = Format$(Val(fgSelect.Text), "000000000000000.00")

    End Select
    fgSelect.Col = fgSelect_arrIndex - 1
    fgSelect.Text = X
Next I

fgSelect_Sort1 = fgSelect_arrIndex - 1: fgSelect_Sort2 = fgSelect_arrIndex - 1
fgSelect_Sort
End Sub

Public Sub fgSelect_Sort_ZCAUDOS0()
Dim I As Integer, K As Integer, X As String, wIndex As Long, blnOk As Boolean, wColor As Long
Dim wMAD As Long, wAmj As String, curX As Currency

For I = 1 To fgSelect.Rows - 1
    fgSelect.Row = I
    fgSelect.Col = 0: X = Trim(fgSelect.Text)
    fgSelect.Col = 3: X = X & Trim(fgSelect.Text)
    fgSelect.Col = 2: X = X & Format$(Val(fgSelect.Text), "000000000000000.00")
    fgSelect.Col = 5: X = X & Trim(fgSelect.Text)
    fgSelect.Col = fgSelect_arrIndex - 1
    fgSelect.Text = X
Next I

fgSelect_Sort1 = fgSelect_arrIndex - 1: fgSelect_Sort2 = fgSelect_arrIndex - 1
fgSelect_SortAD = 5
fgSelect_Sort

ReDim arrCol_CLI(fgSelect.Rows) As String, arrCol_DEV(fgSelect.Rows) As String, arrCol_MTD(fgSelect.Rows) As Currency
ReDim arrCol_NAT(fgSelect.Rows) As String, arrCol_MAD(fgSelect.Rows) As String, arrCol_NUM(fgSelect.Rows) As Long

For I = 1 To fgSelect.Rows - 1
    fgSelect.Row = I
    fgSelect.Col = 0: arrCol_CLI(I) = Trim(fgSelect.Text)
    fgSelect.Col = 3: arrCol_DEV(I) = Trim(fgSelect.Text)
    fgSelect.Col = 2: arrCol_MTD(I) = Val(fgSelect.Text)
    fgSelect.Col = 5: arrCol_NAT(I) = Trim(fgSelect.Text)
    fgSelect.Col = 6: arrCol_NUM(I) = Val(fgSelect.Text)
    fgSelect.Col = 8: arrCol_MAD(I) = Trim(fgSelect.Text)
Next I

wMAD = DSys - 50000

For I = 1 To fgSelect.Rows - 1
    K = I + 1
    blnOk = False
    If arrCol_CLI(I) = arrCol_CLI(K) And arrCol_DEV(I) = arrCol_DEV(K) Then
        If arrCol_NAT(I) <> arrCol_NAT(K) Then
            If arrCol_MTD(I) = arrCol_MTD(K) Then
                blnOk = True: wColor = mColor_G2
            Else
                Call dateJMA_AMJ(arrCol_MAD(I), wAmj)
                If Val(wAmj) < wMAD Then
                    curX = arrCol_MTD(I) * 0.2
                Else
                    curX = arrCol_MTD(I) * 0.1
                End If
                If arrCol_NUM(I) = 3287 Or arrCol_NUM(K) = 3287 Then curX = arrCol_MTD(I) * 0.3  'demande KD 2016-01-29
                If Abs(arrCol_MTD(I) - arrCol_MTD(K)) < curX Then blnOk = True: wColor = mColor_G1
            End If
        End If
    End If
    If blnOk Then
        fgSelect.Row = I
        For K = 0 To 10
            fgSelect.Col = K: fgSelect.CellBackColor = wColor
        Next K
        fgSelect.Row = I + 1
        For K = 0 To 10
            fgSelect.Col = K: fgSelect.CellBackColor = wColor
        Next K
       I = I + 1
    Else
        fgSelect.Row = I
        If arrCol_NAT(I) = "SNE" Then
            wColor = mColor_Y1
        Else
            wColor = RGB(255, 255, 255) 'mColor_Y1
        End If
        For K = 0 To 10
            fgSelect.Col = K: fgSelect.CellBackColor = wColor
        Next K
    End If
    
Next I


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



Private Sub fgSelect_Display()

Dim K As Long

On Error GoTo Error_Handler
currentAction = "fgSelect_Display"

mDATOPECLI = ""
fgSelect.Visible = False
fgSelect_Reset

fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
fgSelect.Row = 0
fgSelect.Col = 2: fgSelect.CellAlignment = 1
fgSelect.Col = 6: fgSelect.CellAlignment = 1
fgSelect.Col = 11: fgSelect.CellAlignment = 1
fgSelect.Col = 12: fgSelect.CellAlignment = 1

fgSelect.Col = 0: fgSelect.CellAlignment = 2
fgSelect.Col = 3: fgSelect.CellAlignment = 2
fgSelect.Col = 5: fgSelect.CellAlignment = 2
fgSelect.Col = 7: fgSelect.CellAlignment = 2
fgSelect.Col = 8: fgSelect.CellAlignment = 2
fgSelect.Col = 9: fgSelect.CellAlignment = 2

Do While Not rsSab.EOF

    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1
    fgSelect_Display_Line
    
    rsSab.MoveNext

Loop

fgSelect.Visible = True

If fgSelect.Rows = 2 And cmdSelect_SQL_K = 1 Then
    fgSelect.Col = 0
    'Call fgDetail_Display(Trim(fgSelect.Text))
End If

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgSelect.Rows - 1): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub
Public Sub fgSelect_Display_Line()
Dim X As String, K As Integer, xCur As Currency, dblTaux As Double, dblMarge As Double
Dim wColor As Long

On Error Resume Next

If cmdSelect_SQL_K = "3" Then
        fgSelect.Col = 0: fgSelect.Text = rsSab("DATOPECLI")
        fgSelect.CellFontBold = True
        fgSelect.Col = 1: fgSelect.Text = Trim(rsSab("CLIENARA1")) & " " & Trim(rsSab("CLIENARA2"))
        fgSelect.CellFontSize = 8
Else
    If mDATOPECLI <> rsSab("DATOPECLI") Then
        fgSelect.Col = 0: fgSelect.Text = rsSab("DATOPECLI")
        fgSelect.CellFontBold = True
        fgSelect.Col = 1: fgSelect.Text = Trim(rsSab("CLIENARA1")) & " " & Trim(rsSab("CLIENARA2"))
        fgSelect.CellFontSize = 8
        For K = 0 To fgSelect.Cols: fgSelect.Col = K: fgSelect.CellBackColor = mColor_G1: Next K
        fgSelect.Rows = fgSelect.Rows + 1
        fgSelect.Row = fgSelect.Rows - 1
        mDATOPECLI = rsSab("DATOPECLI")
        Call lstErr_ChangeLastItem(lstErr, cmdContext, mDATOPECLI): DoEvents
    
    End If
End If
fgSelect.Col = 2: fgSelect.Text = Format(rsSab("DATOPEMNT"), "### ### ### ##0.00")
Select Case rsSab("DATOPEETA")
    Case "03": fgSelect.CellBackColor = RGB(240, 255, 240)
    Case "06": fgSelect.CellBackColor = RGB(228, 228, 228)
    Case "07": fgSelect.CellBackColor = mColor_W1
    Case Else: fgSelect.CellBackColor = vbMagenta
End Select
fgSelect.CellFontSize = 8: fgSelect.CellFontBold = True
fgSelect.CellForeColor = vbBlue
'fgSelect.CellBackColor = RGB(240, 255, 240)
fgSelect.Col = 3: fgSelect.Text = rsSab("DATOPEDEV")
fgSelect.CellBackColor = RGB(240, 255, 240)
fgSelect.CellFontSize = 8: fgSelect.CellFontBold = True
'___________________________________________________________________________________________
X = "select * from " & paramIBM_Library_SAB & ".ZBASTAB0 " _
     & " where BASTABETA  = " & currentZMNURUT0.MNURUTETB _
     & " and BASTABNUM  = 58" _
     & " and BASTABARG  = '" & rsSab("DATOPEOPR") & rsSab("DATOPENAT") & "'"

Set rsSabX = cnsab.Execute(X)
If Not rsSabX.EOF Then
    X = Trim(Mid$(rsSabX("BASTABDON"), 1, 32))
Else
    X = ""
End If

fgSelect.Col = 4: fgSelect.Text = rsSab("DATOPENAT") & " - " & X
fgSelect.CellFontSize = 7
fgSelect.Col = 5: fgSelect.CellBackColor = mColor_Y1
If rsSab("DATOPENAN") = "O" Then
    fgSelect.Text = " Oui"
End If
fgSelect.Col = 6: fgSelect.Text = Format(rsSab("DATOPENUM"), "### ###")
fgSelect.CellForeColor = RGB(128, 128, 128)
fgSelect.Col = 7: fgSelect.Text = " " & dateImp10_S(rsSab("DATOPENEG") + 19000000)
fgSelect.CellFontSize = 8
fgSelect.CellForeColor = RGB(128, 128, 128)
fgSelect.Col = 8: fgSelect.Text = " " & dateImp10_S(rsSab("DATOPEDIS") + 19000000)
fgSelect.CellFontSize = 8
If rsSab("DATOPEECH") > 0 Then
    fgSelect.Col = 9: fgSelect.Text = " " & dateImp10_S(rsSab("DATOPEECH") + 19000000)
Else
    fgSelect.Col = 9: fgSelect.Text = " " & dateImp10_S(rsSab("DATOPEREE") + 19000000)
    fgSelect.CellBackColor = mColor_W0
End If

fgSelect.CellFontSize = 8: fgSelect.CellFontBold = True

'___________________________________________________________________________________________
xCur = 0
X = "select * from " & paramIBM_Library_SAB & ".ZDATVEN0 " _
     & " where DATVENETB  = " & rsSab("DATOPEETB") _
     & " and DATVENAGE  = " & rsSab("DATOPEAGE") _
     & " and DATVENSER  = '" & rsSab("DATOPESER") & "'" _
     & " and DATVENSES  = '" & rsSab("DATOPESES") & "'" _
     & " and DATVENOPR  = '" & rsSab("DATOPEOPR") & "'" _
     & " and DATVENNUM  = " & rsSab("DATOPENUM") _
     & " and DATVENNAT  = '" & rsSab("DATOPENAT") & "'" _
     & " and DATVENTYP  = 'ECH'"

Set rsSabX = cnsab.Execute(X)
Do While Not rsSabX.EOF
    xCur = xCur + rsSabX("DATVENMN2")
    rsSabX.MoveNext

Loop
fgSelect.Col = 12: fgSelect.Text = Format(xCur, "### ### ### ##0.00")
fgSelect.CellFontSize = 8: fgSelect.CellFontBold = True
fgSelect.CellForeColor = vbBlue: fgSelect.CellBackColor = RGB(240, 255, 240)

'____________________________________________________________________________________
X = "select * from " & paramIBM_Library_SAB & ".ZDATCON0 " _
     & " where DATCONETB  = " & rsSab("DATOPEETB") _
     & " and DATCONAGE  = " & rsSab("DATOPEAGE") _
     & " and DATCONSER  = '" & rsSab("DATOPESER") & "'" _
     & " and DATCONSES  = '" & rsSab("DATOPESES") & "'" _
     & " and DATCONOPR  = '" & rsSab("DATOPEOPR") & "'" _
     & " and DATCONNUM  = " & rsSab("DATOPENUM") _
     & " and DATCONNAT  = '" & rsSab("DATOPENAT") & "'" _
     & " and DATCONCON = 1"
     
Set rsSabX = cnsab.Execute(X)

If Not rsSabX.EOF Then
        
    X = ""
    fgSelect.Col = 10: fgSelect.CellForeColor = RGB(128, 0, 0)
    If Trim(rsSabX("DATCONREF")) <> "" Then
        X = Trim(rsSabX("DATCONREF"))
        dblMarge = rsSabX("DATCONMAR")
        Select Case dblMarge
            Case 0: fgSelect.Text = X
            Case Is > 0: fgSelect.Text = X & " + " & Format(dblMarge, "###0.000##")
            Case Is < 0: fgSelect.Text = X & " - " & Format(Abs(dblMarge), "###0.000##")
        End Select
    End If
    
    fgSelect.Col = 11: fgSelect.CellBackColor = mColor_Y1: fgSelect.CellForeColor = RGB(128, 0, 0)
    fgSelect.CellFontSize = 8: fgSelect.CellFontBold = True
    If rsSabX("DATCONTXF") <> 0 Then
        dblTaux = rsSabX("DATCONTXF") + dblMarge
        If dblTaux <> 0 Then
            fgSelect.Text = Format(dblTaux, "###0.000##")
        End If
    End If
End If


'___________________________________________________________________________________________

End Sub

Public Sub Msg_Rcv(Msg As String)
'---------------------------------------------------------
Dim wFct As String

mWindowState = Me.WindowState
If Me.WindowState <> frmUsr_Windowstate Then Me.WindowState = frmUsr_Windowstate

wFct = UCase$(Trim(Mid$(Msg, 1, 12)))
Call BIA_VB_HAB(wFct, arrHab(), cboSelect_SQL)


Select Case wFct
    'Case "@?????":
    Case Else: blnAuto = False: Form_Init

End Select
End Sub



Public Sub fgSelect_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
fgSelect.Visible = False
mRow = fgSelect.Row

If lRow > 0 And lRow < fgSelect.Rows Then
    fgSelect.Row = lRow
    For I = 1 To 0 Step -1
        fgSelect.Col = I: fgSelect.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgSelect.Row = mRow
    If fgSelect.Row > 0 Then
        lRow = fgSelect.Row
        lColor_Old = fgSelect.CellBackColor
        For I = 1 To 0 Step -1
          fgSelect.Col = I: fgSelect.CellBackColor = lColor
        Next I
    End If
End If
fgSelect.LeftCol = fgSelect.FixedCols
fgSelect.Visible = True
End Sub


Private Sub cboSelect_SQL_Click()
cmdSelect_Reset

End Sub


Private Sub cmdPrint_Click()

'If cmdSelect_SQL_K = "1" Then
Me.PopupMenu mnuPrint, vbPopupMenuLeftButton
    

End Sub

Private Sub fgDetail_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim wX As String, xUUMID As String
On Error Resume Next


If y <= fgDetail.RowHeightMin Then
    fgDetail.Visible = False
    Select Case fgDetail.Col
        Case 0: fgDetail_Sort1 = 0: fgDetail_Sort2 = 3: fgDetail_Sort
        Case 1:  fgDetail_Sort1 = 1: fgDetail_Sort2 = 3: fgDetail_Sort
        Case 2: fgDetail_Sort1 = 2: fgDetail_Sort2 = 2: fgDetail_Sort
        Case 3: fgDetail_Sort1 = 3: fgDetail_Sort2 = 3: fgDetail_Sort
        Case 4: fgDetail_Sort1 = 4: fgDetail_Sort2 = 4: fgDetail_Sort
    End Select
    fgDetail.Visible = True
Else
    If fgDetail.Rows > 1 Then
   End If
End If
fgDetail.LeftCol = 0


End Sub




Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim wX As String, xUUMID As String
On Error Resume Next


If y <= fgSelect.RowHeightMin Then
    fgSelect.Visible = False
    Select Case fgSelect.Col
        Case 0: fgSelect_Sort1 = 0: fgSelect_Sort2 = 0: fgSelect_Sort
        Case 1:  fgSelect_Sort1 = 1: fgSelect_Sort2 = 2: fgSelect_Sort
        Case 2: fgSelect_Sort1 = 2: fgSelect_Sort2 = 2: fgSelect_Sort
        Case 3: fgSelect_Sort1 = 3: fgSelect_Sort2 = 3: fgSelect_Sort
    End Select
    fgSelect.Visible = True
Else
    If fgSelect.Rows > 1 Then
        Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
        Select Case cmdSelect_SQL_K
            Case "1"
                fgSelect.Col = 0: wX = Trim(fgSelect.Text)
                'Call fgDetail_Display(wX)
            Case "2"
                fgSelect.Col = 0: wX = Trim(fgSelect.Text)
                'Call fgDetail_Display(wX)
        End Select
        
   End If
End If
fgSelect.LeftCol = 0


End Sub

Private Sub Form_Activate()
Set XForm = Me

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case Is = 13: KeyCode = 0: cmdContext_Return
    Case Is = 27: cmdContext_Quit
'   Case Is = 34: cmdPageNext_Click
'   Case Is = 33: cmdPagePrior_Click
    Case Is = 44: KeyCode = 0: frmElpPrt.prtScreen
End Select

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
blnControl = True

End Sub

Public Sub cmdSelect_Clear()

lstErr.Clear
fgSelect.Visible = False
fgDetail.Visible = False
cmdSelect_Ok.BackColor = vbGreen

End Sub

Public Sub cmdSelect_Reset()
Dim K As Integer
If blnControl Then
    cmdSelect_Clear
    K = InStr(cboSelect_SQL, "-")
    If K > 1 Then
        cmdSelect_SQL_K = Trim(Mid$(cboSelect_SQL, 1, K - 1))
    Else
        cmdSelect_SQL_K = "???"
    End If
    
    fraSelect_Options.Visible = False
    
    Select Case cmdSelect_SQL_K
        Case "1", "3":
            txtSelect_DATOPEDIS.Visible = False: lblSelect_DATOPEDIS.Visible = False
            cmdSelect_Ok.Visible = True: fraSelect_Options.Visible = True
        Case "2":
            txtSelect_DATOPEDIS.Visible = True: lblSelect_DATOPEDIS.Visible = True
            cmdSelect_Ok.Visible = True: fraSelect_Options.Visible = True
    End Select

End If
End Sub


Private Sub cmdSelect_SQL_2()
Dim V, X As String, wAMJ_IBM As Long
Dim xSQL As String, xWhere As String
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_2"
X = Trim(cboSelect_DATOPECLI)
If X <> "" Then
    X = " and DATOPECLI = '" & X & "'"
End If
Call DTPicker_Control(txtSelect_DATOPEDIS, wAMJMin)
wAMJ_IBM = wAMJMin - 19000000
xSQL = "select * from " & paramIBM_Library_SAB & ".ZDATOPE0 , " & paramIBM_Library_SAB & ".ZCLIENA0 " _
     & " where DATOPEDIS <= " & wAMJ_IBM _
     & " And ( DATOPEECH > " & wAMJ_IBM & " or  DATOPEREE > " & wAMJ_IBM & ")" _
     & " and CLIENAETB = DATOPEETB and CLIENACLI = DATOPECLI" _
     & X _
     & " order by DATOPECLI , DATOPENAT , DATOPENUM"
Set rsSab = cnsab.Execute(xSQL)
  

fgSelect_Display

Set rsSab = Nothing

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub cmdSelect_SQL_1()
Dim V, X As String
Dim xSQL As String, xWhere As String
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_1"
X = Trim(cboSelect_DATOPECLI)
If X <> "" Then
    X = " and DATOPECLI = '" & X & "'"
End If

xSQL = "select * from " & paramIBM_Library_SAB & ".ZDATOPE0 , " & paramIBM_Library_SAB & ".ZCLIENA0 " _
     & " where DATOPEETA = '03'" _
     & " and CLIENAETB = DATOPEETB and CLIENACLI = DATOPECLI" _
     & X _
     & " order by DATOPECLI , DATOPENAT , DATOPENUM"
Set rsSab = cnsab.Execute(xSQL)
  

fgSelect_Display

Set rsSab = Nothing

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub


Private Sub cmdSelect_SQL_3()
Dim V, X As String, XCAU As String
Dim xSQL As String, xWhere As String
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_3"
X = Trim(cboSelect_DATOPECLI)
If X <> "" Then
    XCAU = " and CAUDOSBEN = '" & X & "'"
    X = " and DATOPECLI = '" & X & "'"
End If

xSQL = "select * from " & paramIBM_Library_SAB & ".ZDATOPE0 , " & paramIBM_Library_SAB & ".ZCLIENA0 " _
     & " where DATOPEETA = '03' and DATOPEAUT = 'DNA'" _
     & " and CLIENAETB = DATOPEETB and CLIENACLI = DATOPECLI" _
     & X _
     & " order by DATOPECLI , DATOPENAT , DATOPENUM"
Set rsSab = cnsab.Execute(xSQL)
  

fgSelect_Display

fgSelect.Visible = False

xSQL = "select * from " & paramIBM_Library_SAB & ".ZCAUDOS0 , " & paramIBM_Library_SAB & ".ZCLIENA0 " _
     & " where CAUDOSTRA < 4 and CAUDOSCOA = 'SNE'" _
     & " and CLIENAETB = CAUDOSETB and CLIENACLI = CAUDOSBEN" _
     & XCAU _
     & " order by CAUDOSBEN , CAUDOSDOS"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF

    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1
    fgSelect_Display_Line_ZCAUDOS0
    
    rsSab.MoveNext

Loop

fgSelect_Sort_ZCAUDOS0

fgSelect.Visible = True



Set rsSab = Nothing

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

'---------------------------------------------------------
Public Sub arrTag_Set(ByVal B As Boolean)
'---------------------------------------------------------
For I = 0 To arrTagNb: arrTag(I) = B: Next I
End Sub


Public Sub cmdContext_Return()
    If SSTab1.Tab = 0 Then
        cmdSelect_Ok_Click
    Else
        SendKeys "{TAB}"
    End If
End Sub


Public Sub cmdContext_Quit()
lstErr.Clear: lstErr.Height = 200

If txtRTF.Visible Then
    txtRTF.Visible = False
    Exit Sub
End If

If txtFg.Visible Then
    txtFg.Visible = False
    Exit Sub
End If

If fgDetail.Visible Then
    fgDetail.Visible = False
    Exit Sub
End If

If fgSelect.Visible Then
    fgSelect.Visible = False
    Exit Sub
End If


Unload Me

End Sub

Private Sub Form_Load()


mHeight_0 = Me.Height: mWidth_0 = Me.Width: mHeight_2 = 0: mWidth_2 = 0: mWindowState = Me.WindowState
Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)
blnControl = False

End Sub


Private Sub Form_Resize()
If mWindowState <> Me.WindowState Then
    If Me.WindowState = 0 Or Me.WindowState = 2 Then
        Elp_Form_Resize Me, mWindowState, mHeight_0, mWidth_0, mHeight_2, mWidth_2
    End If
End If

End Sub

Private Sub lstErr_Click()
If lstErr.Height > 500 Then
    lstErr.Height = 480
Else
    lstErr.Height = lstErr.ListCount * 200 + 300
End If

End Sub





Private Sub cmdSelect_Ok_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> SAB_DAT_cmdSelect_Ok ........"): DoEvents

If fgSelect.Visible Then cmdSelect_Clear

Select Case cmdSelect_SQL_K
    Case "1": cmdSelect_SQL_1
    Case "2": cmdSelect_SQL_2
    Case "3": cmdSelect_SQL_3
    Case "JPL":
End Select
    
    
Call lstErr_AddItem(lstErr, cmdContext, "< SAB_DAT_cmdSelect_Ok"): DoEvents
lstErr.Height = 480
Me.Enabled = True: Me.MousePointer = 0
If cmdSelect_Ok.Visible Then cmdSelect_Ok.SetFocus
cmdSelect_Ok.BackColor = fgSelect.BackColorFixed
End Sub



Private Sub mnuPrint_Excel_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
Dim X As String
Call lstErr_Clear(lstErr, cmdContext, "> SAB_DAT_Excel ........"): DoEvents

Select Case cmdSelect_SQL_K
    Case "1"
        X = "DAT - Liste des encours DAT au " & dateImp10_S(YBIATAB0_DATE_CPT_JS1)
        Call MSflexGrid_Excel("", "DAT", X, fgSelect, fgSelect.Cols - 1)
    Case "2"
        X = "DAT - Situation DAT au " & dateImp10_S(wAMJMin)
        Call MSflexGrid_Excel("", "DAT", X, fgSelect, fgSelect.Cols - 1)
    Case "3"
        X = "DNA/SNE - Suivi des encours au " & dateImp10_S(YBIATAB0_DATE_CPT_JS1)
        Call MSflexGrid_Excel("", "DAT", X, fgSelect, fgSelect.Cols - 1)
End Select
Call lstErr_AddItem(lstErr, cmdContext, "< SAB_DAT_Excel Terminé"): DoEvents
fgSelect.LeftCol = 0
Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub mnuPrint_Mail_Click()
Dim xObjet As String, xMesg As String
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> SAB_DAT_Mail ........"): DoEvents

Select Case cmdSelect_SQL_K
    Case "1"
        xObjet = "DAT - Liste des encours DAT au " & dateImp10_S(YBIATAB0_DATE_CPT_JS1)
        xMesg = "<span style='font-size:10.0pt;font-family:Calibri'>" _
         & xObjet
    
        Call MSFlexGrid_SendMail(currentSSIWINMAIL, "DAT", xObjet, xMesg, fgSelect, fgSelect.Cols - 1)
    Case "2"
        xObjet = "DAT - situation DAT au " & dateImp10_S(wAMJMin)
        xMesg = "<span style='font-size:10.0pt;font-family:Calibri'>" _
         & xObjet
    
        Call MSFlexGrid_SendMail(currentSSIWINMAIL, "DAT", xObjet, xMesg, fgSelect, fgSelect.Cols - 1)
    Case "3"
        xObjet = "DNA/SNE - Suivi des encours au " & dateImp10_S(YBIATAB0_DATE_CPT_JS1)
        xMesg = "<span style='font-size:10.0pt;font-family:Calibri'>" _
         & xObjet
    
        Call MSFlexGrid_SendMail(currentSSIWINMAIL, "DAT", xObjet, xMesg, fgSelect, fgSelect.Cols - 1)
End Select

Call lstErr_Clear(lstErr, cmdContext, "< SAB_DAT_Mail Terminé"): DoEvents
fgSelect.LeftCol = 0
Me.Enabled = True: Me.MousePointer = 0

End Sub



Public Sub fgSelect_Display_Line_ZCAUDOS0()
Dim X As String, K As Integer, xCur As Currency, dblTaux As Double, dblMarge As Double
Dim wColor As Long
Dim newMontant As Double

On Error Resume Next

fgSelect.Col = 0: fgSelect.Text = rsSab("CAUDOSBEN")
fgSelect.CellFontBold = True
fgSelect.Col = 1: fgSelect.Text = Trim(rsSab("CLIENARA1")) & " " & Trim(rsSab("CLIENARA2"))
fgSelect.CellFontSize = 8
newMontant = 0
'Voir si le montant de l'encours a été bougé (encours restant)
If retourne_encours_restant(rsSab("CAUDOSDOS"), rsSab("CAUDOSBEN"), rsSab("CAUDOSETB"), newMontant) = False Then
    newMontant = CDbl(rsSab("CAUDOSMNT"))
End If
'                                                           '
fgSelect.Col = 2: fgSelect.Text = Format(newMontant, "### ### ### ##0.00")
fgSelect.Col = 3: fgSelect.Text = rsSab("CAUDOSDEV")
fgSelect.Col = 4: fgSelect.Text = rsSab("CAUDOSCAU")
fgSelect.Col = 5: fgSelect.Text = " " & Trim(rsSab("CAUDOSCOA"))

fgSelect.Col = 6: fgSelect.Text = Format(rsSab("CAUDOSDOS"), "### ###")
fgSelect.CellForeColor = RGB(128, 128, 128)
fgSelect.Col = 7: fgSelect.Text = " " & dateImp10_S(rsSab("CAUDOSCRE") + 19000000)
fgSelect.CellFontSize = 8
fgSelect.CellForeColor = RGB(128, 128, 128)
fgSelect.Col = 8: fgSelect.Text = " " & dateImp10_S(rsSab("CAUDOSDEB") + 19000000)
fgSelect.CellFontSize = 8
If rsSab("CAUDOSNEW") > 0 Then
    fgSelect.Col = 9: fgSelect.Text = " " & dateImp10_S(rsSab("CAUDOSNEW") + 19000000)
End If
fgSelect.Col = 10: fgSelect.Text = rsSab("CAUDOSACT")

End Sub
