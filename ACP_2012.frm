VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmACP 
   AutoRedraw      =   -1  'True
   Caption         =   "JPL"
   ClientHeight    =   10305
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   13530
   BeginProperty Font 
      Name            =   "@Arial Unicode MS"
      Size            =   7.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ACP_2012.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   10305
   ScaleWidth      =   13530
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      TabPicture(0)   =   "ACP_2012.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraSelect"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "ACP_2012.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtFg"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "txtRTF"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "ACP_2012.frx":0342
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
         Text            =   "ACP_2012.frx":035E
         Top             =   1155
         Visible         =   0   'False
         Width           =   5775
      End
      Begin VB.Frame fraSelect 
         BackColor       =   &H00E0E0E0&
         Height          =   9630
         Left            =   -135
         TabIndex        =   4
         Top             =   495
         Width           =   13425
         Begin MSFlexGridLib.MSFlexGrid fgDetail 
            Height          =   7680
            Left            =   5415
            TabIndex        =   13
            Top             =   1530
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
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   11820
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   705
            Width           =   1335
         End
         Begin VB.ComboBox cboSelect_SQL 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
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
            Left            =   120
            TabIndex        =   5
            Top             =   120
            Visible         =   0   'False
            Width           =   9375
            Begin VB.ComboBox cboSelect_App 
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1560
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   12
               Top             =   450
               Visible         =   0   'False
               Width           =   1860
            End
            Begin VB.Label lblSelect_App 
               BackColor       =   &H00F0FFFF&
               Caption         =   "Application"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   225
               TabIndex        =   11
               Top             =   480
               Visible         =   0   'False
               Width           =   1155
            End
         End
         Begin MSFlexGridLib.MSFlexGrid fgSelect 
            Height          =   7710
            Left            =   120
            TabIndex        =   10
            Top             =   1425
            Visible         =   0   'False
            Width           =   5235
            _ExtentX        =   9234
            _ExtentY        =   13600
            _Version        =   393216
            FixedCols       =   0
            RowHeightMin    =   400
            BackColor       =   16777215
            ForeColor       =   16711680
            BackColorFixed  =   8421376
            ForeColorFixed  =   16777215
            BackColorBkg    =   16777215
            GridColor       =   10526720
            GridColorFixed  =   10526720
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            FormatString    =   "<Code                   |<Intitulé                                                                    "
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
         TextRTF         =   $"ACP_2012.frx":0366
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
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      Picture         =   "ACP_2012.frx":03E6
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
   End
End
Attribute VB_Name = "frmACP"
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
Dim BIA_VB_Habilitations_Aut As typeAuthorization
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

Dim arrHab(20) As Boolean

Dim HeightOfLine As Long, LinesOfText As Long

Dim txtRTF_prtForeColor_Header As Long

Dim rsSabX As New ADODB.Recordset

Dim wFile As String
Dim appExcel As Excel.Application 'Application Excel
Dim wbExcel As Excel.Workbook 'Classeur Excel
Dim wsExcel As Excel.Worksheet 'Feuille Excel
Dim mXls1_Row As Long, mXls1_Cols As Integer, mXls1_File As Integer
Dim mXls2_Cols As Integer, mXls2_Row As Integer

Dim arrNature_Code() As String, arrNature_Lib() As String, arrNature_Nb As Integer
Dim arrCLIENACAT_Code() As String, arrCLIENACAT_Lib() As String, arrCLIENACAT_Nb As Integer
Dim arrCLIENAECO_Code() As String, arrCLIENAECO_Lib() As String, arrCLIENAECO_Nb As Integer
Dim arrPays() As typePays, arrPays_Nb As Integer
Dim wAMJMin_IBM As Long
Public Sub cmdPrint_Excel()
On Error GoTo Error_Handler
Dim xSQL As String
Dim X As String, wFilex As String
Dim blnCALCS As Boolean

On Error GoTo Error_Handler
'===================================================================================
'______________________________________________'
X = paramServer("\\CPT_Archive\")
wAMJMin = YBIATAB0_DATE_CPT_J
wAMJMin_IBM = wAMJMin - 19000000

blnCALCS = False
If Dir(X & "CALCS.jpl") <> "" Then blnCALCS = True
blnCALCS = True

If X = "" Then X = "C:\Temp\"
If mId$(X, Len(X), 1) <> "\" Then X = X & "\"

    
wFile = X & Trim("ACP" & " -" & DSYS_Time & mXls1_File & ".xlsx")
'______________________________________________
If Not blnAuto Then
    X = InputBox("par défaut : " _
        & vbCrLf & "     =========================" & vbCrLf & wFile _
        & vbCrLf & "     =========================", "ACP : nom du fichier d'exportation", wFile)
    If Trim(X) = "" Then mXls1_File = mXls1_File - 1: Exit Sub
    wFilex = Trim(X)
    '______________________________________________
    If wFile <> wFilex Then
        wFile = wFilex
    End If
End If
'_________________________________________


If Dir(wFile) <> "" Then msFileSystem.DeleteFile wFile

'=========================================================================================
Call lstErr_AddItem(lstErr, cmdContext, "Fichier excel.... : "): DoEvents

Set appExcel = CreateObject("Excel.Application")
appExcel.Workbooks.Add
Set wbExcel = appExcel.ActiveWorkbook
With wbExcel
    .Title = "ACP"
    .Subject = "ACP"
End With

'__________________________________________________________________________________

'appExcel.Worksheets.Add

Set wsExcel = wbExcel.Sheets(1): wsExcel.Name = "ACP "
Set wsExcel = wbExcel.Sheets(1)

With wsExcel.Cells
    .Borders.Weight = xlWide
    .Borders.Color = RGB(255, 255, 153)
    .Borders(xlInsideHorizontal).Weight = xlThin ' xlMedium
    .Borders(xlInsideHorizontal).Color = RGB(128, 128, 255)
    .Borders(xlInsideVertical).Weight = xlThin
    .Borders(xlInsideVertical).Color = RGB(128, 128, 255)
    .HorizontalAlignment = Excel.xlHAlignRight
    .WrapText = False ' True
    .Font.Size = 8
    .Font.Name = "Calibri"
    .RowHeight = 17
End With

wsExcel.PageSetup.Orientation = vbPRORLandscape
wsExcel.PageSetup.HeaderMargin = 5
wsExcel.PageSetup.Zoom = 80

wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14En-cours TC, DAT, Crédits, arrêté au " & dateImp10(wAMJMin) _
                                 & vbCr

wsExcel.PageSetup.CenterHorizontally = True


wsExcel.PageSetup.PrintTitleRows = "$A1:$K1"
wsExcel.PageSetup.LeftFooter = "&F - &B&A"
wsExcel.PageSetup.RightFooter = "&P / &N"

mXls1_Row = 1

Select Case SSTab1.Tab
    Case 0:
        Select Case cmdSelect_SQL_K
            Case "1":   appExcel.Worksheets.Add
            
                        Set wsExcel = wbExcel.Sheets(1)
                        Call cmdPrint_Excel_BIC
                        Set wsExcel = wbExcel.Sheets(2)
                        Call cmdPrint_Excel_ZTREOPE0
                        Call cmdPrint_Excel_ZCHGOPE0
                        Set wsExcel = wbExcel.Sheets(3)
                        Call cmdPrint_Excel_ZDATOPE0
                        Set wsExcel = wbExcel.Sheets(4)
                        Call cmdPrint_Excel_ZCREPRE0
      End Select
        

End Select
'======================================================================================================

Exit_sub:
'__________________________________________________________________________________
Set rsSab = Nothing


wbExcel.SaveAs wFile

wbExcel.Close

'____________________________________________________________________________________
appExcel.Quit

Set rsSab = Nothing

Set wsExcel = Nothing
Set wbExcel = Nothing
Set appExcel = Nothing
Call lstErr_AddItem(lstErr, cmdContext, "< Exportation terminée"): DoEvents


'_____________________________
Exit Sub

Error_Handler:
    If Not blnCALCS Then
        X = "C:\Temp\"
        Resume Next
    End If
    MsgBox Error, vbCritical, Me.Name
    Call lstErr_AddItem(lstErr, cmdContext, "< Exportation terminée"): DoEvents
    
    wbExcel.SaveAs wFile
    wbExcel.Close
    appExcel.Quit

End Sub

Public Sub cmdPrint_Excel_BIC()
Dim xSQL As String, X As String, K As Long, K2 As Long
On Error GoTo Error_Handler

'===================================================================================

wsExcel.Name = "BIC"

With wsExcel.Cells
    .Borders.Weight = xlWide
    .Borders.Color = RGB(255, 255, 153)
    .Borders(xlInsideHorizontal).Weight = xlThin ' xlMedium
    .Borders(xlInsideHorizontal).Color = RGB(128, 128, 255)
    .Borders(xlInsideVertical).Weight = xlThin
    .Borders(xlInsideVertical).Color = RGB(128, 128, 255)
    .HorizontalAlignment = Excel.xlHAlignLeft
    .WrapText = False ' True
    .Font.Size = 9
    .Font.Name = "Calibri"
    .RowHeight = 17
End With

wsExcel.PageSetup.Orientation = vbPRORLandscape
wsExcel.PageSetup.HeaderMargin = 5
wsExcel.PageSetup.Zoom = 80

wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14Correspondance BIC - racine , arrêté au " & dateImp10(wAMJMin) _
                                 & vbCr

wsExcel.PageSetup.CenterHorizontally = True


wsExcel.PageSetup.PrintTitleRows = "$A1:$K1"
wsExcel.PageSetup.LeftFooter = "&F - &B&A"
wsExcel.PageSetup.RightFooter = "&P / &N"

mXls1_Row = 1
mXls1_Cols = 3

wsExcel.Columns(1).ColumnWidth = 12: wsExcel.Cells(1, 1) = "BIC "
wsExcel.Columns(2).ColumnWidth = 12: wsExcel.Cells(1, 2) = "Racine"
wsExcel.Columns(3).ColumnWidth = 60: wsExcel.Cells(1, 3) = "Intitulé"


For K = 1 To mXls1_Cols
    wsExcel.Cells(1, K).Interior.Color = mColor_GB
    wsExcel.Cells(1, K).Font.Color = vbWhite
Next



xSQL = "select * from " & paramIBM_Library_SAB & ".ZADRESS0 , " & paramIBM_Library_SAB & ".ZCLIENA0" _
     & " where ADRESSTYP = 4  and substring(ADRESSNUM,2,7) = CLIENACLI" _
     & " order by ADRESSRA12  "
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    Call lstErr_ChangeLastItem(lstErr, cmdContext, rsSab("ADRESSRA12")): DoEvents
     mXls1_Row = mXls1_Row + 1

     wsExcel.Cells(mXls1_Row, 1) = Replace(rsSab("ADRESSRA12"), "XXX", "")
     wsExcel.Cells(mXls1_Row, 2) = rsSab("CLIENACLI")
     wsExcel.Cells(mXls1_Row, 3) = Trim(rsSab("CLIENARA1")) & Trim(rsSab("CLIENARA2"))
        
'____________________________________________________________________________________

    rsSab.MoveNext
Loop




Exit Sub
'======================================================================================================

Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
Exit_sub:


End Sub


Public Sub cmdPrint_Excel_ZTREOPE0()
Dim xSQL As String, X As String, K As Long, K2 As Long, xECO As String
Dim curX As Currency, wDev As String, dblCours As Double, dblX As Double
On Error GoTo Error_Handler

'===================================================================================

wsExcel.Name = "FOTC"
cmdPrint_Excel_Init ("En-cours FOTC")

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIASTO0 " _
     & " where YSTOAPP = 'TRE'" _
     & " order by YSTOCLI , YSTOOPE , YSTONAT , YSTONUM "
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
     mXls1_Row = mXls1_Row + 1
     X = rsSab("YSTOSER") & " " & rsSab("YSTOSSE") & " " & rsSab("YSTOOPE") & " " & rsSab("YSTONAT") & "    " & Format(rsSab("YSTONUM"), "### ###")
    Call lstErr_ChangeLastItem(lstErr, cmdContext, X): DoEvents

     wsExcel.Cells(mXls1_Row, 1) = mId$(rsSab("YSTOPCI"), 1, 5)
     wsExcel.Cells(mXls1_Row, 3) = X
     
     wDev = rsSab("YSTODEV")
     wsExcel.Cells(mXls1_Row, 5) = wDev
     wsExcel.Cells(mXls1_Row, 8) = dateImp10(rsSab("YSTODEB"))
     wsExcel.Cells(mXls1_Row, 9) = dateImp10(rsSab("YSTOFIN"))
     
    Call cmdPrint_Excel_ZCLIENA0(Format(rsSab("YSTOCLI"), "0000000"), mId$(rsSab("YSTOPCI"), 1, 5), Trim(rsSab("YSTOOPE") & rsSab("YSTONAT")))

     If rsSab("YSTOOPE") = "PRE" Then
        curX = rsSab("YSTOMON")
    Else
        curX = -rsSab("YSTOMON")
    End If
    wsExcel.Cells(mXls1_Row, 12) = curX

    If wDev <> "EUR" Then
        Call sqlYBIATAB0_Read("PDC", wDev, wAMJMin, X)
        If IsNumeric(mId$(X, 9, 15)) Then
            dblCours = CDbl(mId$(X, 9, 15) / 1000000000)
            If dblCours <> 0 Then curX = Round(curX / dblCours, 2)
        Else
            curX = 0
        End If
    End If
     wsExcel.Cells(mXls1_Row, 13) = curX

     wsExcel.Cells(mXls1_Row, 10) = rsSab("YSTOCTX")
     dblX = rsSab("YSTOTAU")
     If dblX <> 0 Then wsExcel.Cells(mXls1_Row, 11) = dblX
     
    xSQL = "select * from " & paramIBM_Library_SAB & ".ZTREOPE0 " _
         & " where TREOPEETB = 1 and  TREOPEAGE = 1" _
         & " and TREOPESER = '" & rsSab("YSTOSER") & "' and  TREOPESES = '" & rsSab("YSTOSSE") & "'" _
          & " and TREOPEOPR = '" & rsSab("YSTOOPE") & "' and  TREOPENUM = " & rsSab("YSTONUM") _
         & " and TREOPENAT = '" & Trim(rsSab("YSTONAT")) & "'"
    Set rsSabX = cnsab.Execute(xSQL)
    
    If Not rsSabX.EOF Then
         If rsSabX("TREOPEOPR") = "PRE" Then
            curX = rsSabX("TREOPEMNT")
        Else
            curX = -rsSabX("TREOPEMNT")
        End If
        wsExcel.Cells(mXls1_Row, 6) = curX
    
        If wDev <> "EUR" Then
            Call sqlYBIATAB0_Read("PDC", wDev, rsSab("YSTODEB"), X)
            If IsNumeric(mId$(X, 9, 15)) Then
                dblCours = CDbl(mId$(X, 9, 15) / 1000000000)
                If dblCours <> 0 Then curX = Round(curX / dblCours, 2)
            Else
                curX = 0
            End If
        End If
         wsExcel.Cells(mXls1_Row, 7) = curX
     
     End If
     
        xSQL = "select * from " & paramIBM_Library_SAB & ".ZTRECON0 " _
         & " where TRECONETB = 1 and  TRECONAGE = 1" _
         & " and TRECONSER = '" & rsSab("YSTOSER") & "' and  TRECONSES = '" & rsSab("YSTOSSE") & "'" _
          & " and TRECONOPR = '" & rsSab("YSTOOPE") & "' and  TRECONNUM = " & rsSab("YSTONUM") _
         & " and TRECONNAT = '" & Trim(rsSab("YSTONAT")) & "'" _
         & " order by TRECONCON"
    Set rsSabX = cnsab.Execute(xSQL)

    Do While Not rsSabX.EOF
        If wAMJMin_IBM >= rsSabX("TRECONDEB") And wAMJMin_IBM <= rsSabX("TRECONFIN") Then
            If rsSabX("TRECONTXP") <> 0 Then
                dblX = rsSabX("TRECONTXP")
            Else
                dblX = rsSabX("TRECONTXE")
            End If
            If dblX <> 0 Then wsExcel.Cells(mXls1_Row, 11) = dblX
            
            X = ""
            If Trim(rsSabX("TRECONPRE")) <> "" Then
                X = Trim(rsSabX("TRECONPRE"))
                dblX = rsSabX("TRECONMPR")
            Else
                X = Trim(rsSabX("TRECONEMP"))
                dblX = rsSabX("TRECONMEM")
            End If
            If X <> "" Then
                Select Case dblX
                    Case 0: wsExcel.Cells(mXls1_Row, 10) = X
                    Case Is > 0: wsExcel.Cells(mXls1_Row, 10) = X & " + " & Format(dblX, "###0.000##")
                    Case Is < 0: wsExcel.Cells(mXls1_Row, 10) = X & " - " & Format(Abs(dblX), "###0.000##")
                End Select
            End If
       
            Exit Do
        End If
        rsSabX.MoveNext
    Loop
     
'____________________________________________________________________________________

    rsSab.MoveNext
Loop




Exit Sub
'======================================================================================================

Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
Exit_sub:


End Sub
Public Sub cmdPrint_Excel_ZDATOPE0()
Dim xSQL As String, X As String, K As Long, K2 As Long, xNAT As String
Dim curX As Currency, wDev As String, dblCours As Double, dblX As Double
On Error GoTo Error_Handler

'===================================================================================

wsExcel.Name = "DAT"
Call cmdPrint_Excel_Init("En-cours DAT")

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIASTO0 " _
     & " where YSTOAPP = 'DAT'" _
     & " order by YSTOCLI , YSTOOPE , YSTONAT , YSTONUM "
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
     mXls1_Row = mXls1_Row + 1
     xNAT = mId$(rsSab("YSTONAT"), 1, 3)
     X = rsSab("YSTOSER") & " " & rsSab("YSTOSSE") & " " & rsSab("YSTOOPE") & " " & xNAT & "    " & Format(rsSab("YSTONUM"), "### ###")
    Call lstErr_ChangeLastItem(lstErr, cmdContext, X): DoEvents

     wsExcel.Cells(mXls1_Row, 1) = mId$(rsSab("YSTOPCI"), 1, 5)
     wsExcel.Cells(mXls1_Row, 3) = X
     
     wDev = rsSab("YSTODEV")
     wsExcel.Cells(mXls1_Row, 5) = wDev
     wsExcel.Cells(mXls1_Row, 8) = dateImp10(rsSab("YSTODEB"))
     wsExcel.Cells(mXls1_Row, 9) = dateImp10(rsSab("YSTOFIN"))
     
    Call cmdPrint_Excel_ZCLIENA0(Format(rsSab("YSTOCLI"), "0000000"), mId$(rsSab("YSTOPCI"), 1, 5), Trim(rsSab("YSTOOPE") & xNAT))

    curX = rsSab("YSTOMON")
    wsExcel.Cells(mXls1_Row, 12) = curX

    If wDev <> "EUR" Then
        Call sqlYBIATAB0_Read("PDC", wDev, wAMJMin, X)
        If IsNumeric(mId$(X, 9, 15)) Then
            dblCours = CDbl(mId$(X, 9, 15) / 1000000000)
            If dblCours <> 0 Then curX = Round(curX / dblCours, 2)
        Else
            curX = 0
        End If
    End If
     wsExcel.Cells(mXls1_Row, 13) = curX

     wsExcel.Cells(mXls1_Row, 10) = rsSab("YSTOCTX")
     dblX = rsSab("YSTOTAU")
     If dblX <> 0 Then wsExcel.Cells(mXls1_Row, 11) = dblX
     
'____________________________________________________________________________________
    xSQL = "select * from " & paramIBM_Library_SAB & ".ZDATOPE0 " _
         & " where DATOPEETB = 1 and  DATOPEAGE = 1" _
         & " and DATOPESER = '" & rsSab("YSTOSER") & "' and  DATOPESES = '" & rsSab("YSTOSSE") & "'" _
          & " and DATOPEOPR = '" & rsSab("YSTOOPE") & "' and  DATOPENUM = " & rsSab("YSTONUM") _
         & " and DATOPENAT = '" & xNAT & "'"
    Set rsSabX = cnsab.Execute(xSQL)
    
    If Not rsSabX.EOF Then
        curX = rsSabX("DATOPEMNT")
        wsExcel.Cells(mXls1_Row, 6) = curX
    
    If wDev <> "EUR" Then
        Call sqlYBIATAB0_Read("PDC", wDev, rsSab("YSTODEB"), X)
        If IsNumeric(mId$(X, 9, 15)) Then
            dblCours = CDbl(mId$(X, 9, 15) / 1000000000)
            If dblCours <> 0 Then curX = Round(curX / dblCours, 2)
        Else
            curX = 0
        End If
    End If
         wsExcel.Cells(mXls1_Row, 7) = curX
     
     End If
'____________________________________________________________________________________
    xSQL = "select * from " & paramIBM_Library_SAB & ".ZDATCON0 " _
         & " where DATCONETB = 1 and  DATCONAGE = 1" _
         & " and DATCONSER = '" & rsSab("YSTOSER") & "' and  DATCONSES = '" & rsSab("YSTOSSE") & "'" _
          & " and DATCONOPR = '" & rsSab("YSTOOPE") & "' and  DATCONNUM = " & rsSab("YSTONUM") _
         & " and DATCONNAT = '" & xNAT & "'" _
         & " order by DATCONCON"
    Set rsSabX = cnsab.Execute(xSQL)

    Do While Not rsSabX.EOF
        If wAMJMin_IBM >= rsSabX("DATCONDEB") And wAMJMin_IBM <= rsSabX("DATCONFIN") Then
            If rsSabX("DATCONTXF") <> 0 Then
                dblX = rsSabX("DATCONTXF")
                If dblX <> 0 Then wsExcel.Cells(mXls1_Row, 11) = dblX
            End If
            
            X = ""
            If Trim(rsSabX("DATCONREF")) <> "" Then
                X = Trim(rsSabX("DATCONREF"))
                dblX = rsSabX("DATCONMAR")
                Select Case dblX
                    Case 0: wsExcel.Cells(mXls1_Row, 10) = X
                    Case Is > 0: wsExcel.Cells(mXls1_Row, 10) = X & " + " & Format(dblX, "###0.000##")
                    Case Is < 0: wsExcel.Cells(mXls1_Row, 10) = X & " - " & Format(Abs(dblX), "###0.000##")
                End Select
            End If
       
            Exit Do
        End If
        rsSabX.MoveNext
    Loop
     
'____________________________________________________________________________________
    

    rsSab.MoveNext
Loop




Exit Sub
'======================================================================================================

Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
Exit_sub:


End Sub
Public Sub cmdPrint_Excel_ZCREPRE0()
Dim xSQL As String, X As String, K As Long, K2 As Long, xNAT As String, wAMJDEB As Long, wAMJFIN As Long, wCREPREPLA As Long
Dim curX As Currency, wDev As String, dblCours As Double, dblX As Double, wCREPREDIC As Long, wCREPREDAE As Long, wCREPREDPE As Long
Dim blnActif As Boolean, wPCI As String, wCli As String
On Error GoTo Error_Handler

'===================================================================================

wsExcel.Name = "Crédit"
Call cmdPrint_Excel_Init("En-cours Crédits")

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIASTO0 " _
     & " where YSTOAPP = 'CRE'" _
     & " order by YSTOCLI , YSTOOPE , YSTONAT , YSTONUM "
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
     mXls1_Row = mXls1_Row + 1
     blnActif = True

     
     xNAT = Trim(rsSab("YSTONAT"))
     wDev = rsSab("YSTODEV")
     wAMJDEB = rsSab("YSTODEB")
     wAMJFIN = rsSab("YSTOFIN")
    xSQL = "select * from " & paramIBM_Library_SAB & ".ZCREDOS0 " _
         & " where CREDOSETA = 1 and  CREDOSAGE = 1" _
         & " and CREDOSSER = '" & rsSab("YSTOSER") & "' and  CREDOSSSE = '" & rsSab("YSTOSSE") & "'" _
          & " and  CREDOSDOS = " & rsSab("YSTONUM")
          
    Set rsSabX = cnsab.Execute(xSQL)
    
    If Not rsSabX.EOF Then
        wAMJDEB = rsSabX("CREDOSDDE") + 19000000
       ' wAMJFIN = rsSabX("CREDOSDFI") + 19000000
   
    End If
    
    
    xSQL = "select * from " & paramIBM_Library_SAB & ".ZCREPRE0 " _
         & " where CREPREETA = 1 and  CREPREAGE = 1" _
         & " and CREPRESCE = '" & rsSab("YSTOSER") & "' and  CREPRESSE = '" & rsSab("YSTOSSE") & "'" _
          & " and  CREPREDOS = " & rsSab("YSTONUM") & " and  CREPREPRE = " & rsSab("YSTOSEQ")
          
    Set rsSabX = cnsab.Execute(xSQL)
    
    If Not rsSabX.EOF Then
        If rsSabX("CREPRECTA") <> 7 Then
            blnActif = False
            If rsSabX("CREPRECTA") <> 8 Then Call MsgBox("CREPRESTA  =  " & rsSabX("CREPRECTA") & " dossier : " & rsSab("YSTONUM"), vbInformation, "ACP")
        End If
        
        xNAT = Trim(rsSabX("CREPRENAT"))
        wDev = rsSabX("CREPREDEV")
        wCREPREPLA = rsSabX("CREPREPLA")
        wCREPREDIC = rsSabX("CREPREDIC")
        wCREPREDAE = rsSabX("CREPREDAE")
        wCREPREDPE = rsSabX("CREPREDPE")
        ''''wAMJDEB = rsSabX("CREPREOUV") + 19000000
        curX = rsSabX("CREPREMON")
        wsExcel.Cells(mXls1_Row, 6) = curX
    
        If wDev <> "EUR" Then
            Call sqlYBIATAB0_Read("PDC", wDev, CStr(wAMJDEB), X)
            If IsNumeric(mId$(X, 9, 15)) Then
                dblCours = CDbl(mId$(X, 9, 15) / 1000000000)
                If dblCours <> 0 Then curX = Round(curX / dblCours, 2)
            Else
            '________________________________________________________________________________
                dblCours = ZBASTAB0_Cours(wDev, CStr(wAMJDEB))
                If dblCours > 0 Then
                    curX = Round(curX / dblCours, 2)
                Else
                    curX = 0
                End If
            End If
'________________________________________________________________________________

        End If
         wsExcel.Cells(mXls1_Row, 7) = curX
     End If
'____________________________________________________________________________________

     X = rsSab("YSTOSER") & " " & rsSab("YSTOSSE") & " " & rsSab("YSTOOPE") & " " & xNAT & "    " & Format(rsSab("YSTONUM"), "### ###") & "-" & rsSab("YSTOSEQ")
    Call lstErr_ChangeLastItem(lstErr, cmdContext, X): DoEvents

     wsExcel.Cells(mXls1_Row, 3) = X
     
     wCli = Format(rsSab("YSTOCLI"), "0000000")
     If Trim(rsSab("YSTOPCI")) <> "" Then
        wPCI = mId$(rsSab("YSTOPCI"), 1, 5)
    Else
        Select Case wCli & xNAT
            Case "0050411PHB": wPCI = "13121"
            Case "0050713PKD": wPCI = "29115"
            Case "0012327PDT": wPCI = "39115"
            Case "0050546PBD": wPCI = "19115"
            Case "0050713PKI": wPCI = "29715"
            Case "0011520PEG": wPCI = "20111"
            Case "0050546PBI": wPCI = "19715"
            Case "0050545PBD": wPCI = "19115"
            Case "0050545PBI": wPCI = "19715"
        End Select
    End If
     wsExcel.Cells(mXls1_Row, 1) = wPCI
     
    Call cmdPrint_Excel_ZCLIENA0(wCli, wPCI, Trim(rsSab("YSTOOPE") & xNAT))
     wsExcel.Cells(mXls1_Row, 5) = wDev
     wsExcel.Cells(mXls1_Row, 8) = dateImp10(wAMJDEB)
     wsExcel.Cells(mXls1_Row, 9) = dateImp10(wAMJFIN)

    curX = rsSab("YSTOMON")
    wsExcel.Cells(mXls1_Row, 12) = curX

    If wDev <> "EUR" Then
        Call sqlYBIATAB0_Read("PDC", wDev, wAMJMin, X)
        If IsNumeric(mId$(X, 9, 15)) Then
            dblCours = CDbl(mId$(X, 9, 15) / 1000000000)
            If dblCours <> 0 Then curX = Round(curX / dblCours, 2)
        Else
            curX = 0
        End If
    End If
     wsExcel.Cells(mXls1_Row, 13) = curX

     
'____________________________________________________________________________________

     wsExcel.Cells(mXls1_Row, 10) = rsSab("YSTOCTX")
     dblX = rsSab("YSTOTAU")
     If dblX <> 0 Then wsExcel.Cells(mXls1_Row, 11) = dblX
    xSQL = "select * from " & paramIBM_Library_SAB & ".ZCREPLA0 " _
         & " where CREPLAETA = 1 and  CREPLAAGE = 1" _
         & " and CREPLASER = '" & rsSab("YSTOSER") & "' and  CREPLASSE = '" & rsSab("YSTOSSE") & "'" _
         & " and  CREPLADOS = " & rsSab("YSTONUM") & " and  CREPLAPRE = " & rsSab("YSTOSEQ") _
         & " and  CREPLAPLA = " & wCREPREPLA
          
    Set rsSabX = cnsab.Execute(xSQL)
    
    If Not rsSabX.EOF Then
                dblX = rsSabX("CREPLATAF") + rsSabX("CREPLAMAR")
                If dblX <> 0 Then wsExcel.Cells(mXls1_Row, 11) = dblX
            
            X = ""
            If Trim(rsSabX("CREPLARTA")) <> "" Then
                X = Trim(rsSabX("CREPLARTA"))
            Else
                X = "[ " & Format(rsSabX("CREPLATAF"), "###0.000##") & " ]"
            End If
                dblX = rsSabX("CREPLAMAR")
                Select Case dblX
                    Case 0: wsExcel.Cells(mXls1_Row, 10) = X
                    Case Is > 0: wsExcel.Cells(mXls1_Row, 10) = X & " + " & Format(dblX, "###0.000##")
                    Case Is < 0: wsExcel.Cells(mXls1_Row, 10) = X & " - " & Format(Abs(dblX), "###0.000##")
                End Select
   
    End If
    

'____________________________________________________________________________________
'If rsSab("YSTONUM") = 652 Then
'    Debug.Print
'End If
'____________________________________________________________________________________

    xSQL = "select * from " & paramIBM_Library_SAB & ".ZCREBIS0_BIS0001" _
         & " where CREBISETA = 1 and  CREBISAGE = 1" _
         & " and CREBISSER = '" & rsSab("YSTOSER") & "' and  CREBISSSE = '" & rsSab("YSTOSSE") & "'" _
         & " and  CREBISDOS = " & rsSab("YSTONUM") & " and  CREBISPRE = " & rsSab("YSTOSEQ") _
         & " and  CREBISTYP in ('02' , '03')" '_
         '& " and  CREBISDEB = " & wCREPREDAE & " and  CREBISFIN = " & wCREPREDPE
    Set rsSabX = cnsab.Execute(xSQL)
    
    Do While Not rsSabX.EOF
        If wAMJMin_IBM >= rsSabX("CREBISDEB") And wAMJMin_IBM <= rsSabX("CREBISFIN") Then
            dblX = rsSabX("CREBISTAU")
            If dblX <> 0 Then wsExcel.Cells(mXls1_Row, 11) = dblX: Exit Do
        End If
     rsSabX.MoveNext
   Loop
    
     
'____________________________________________________________________________________
    
    
     If Not blnActif Then mXls1_Row = mXls1_Row - 1

    rsSab.MoveNext
Loop




Exit Sub
'======================================================================================================

Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
Exit_sub:


End Sub


Public Sub cmdPrint_Excel_ZCHGOPE0()
Dim xSQL As String, X As String, K As Long, K2 As Long, xECO As String
Dim curX As Currency, wDev As String, dblCours As Double, dblX As Double, wMTD As Currency, wMTE As Currency
On Error GoTo Error_Handler

'===================================================================================


xSQL = "select * from " & paramIBM_Library_SAB & ".ZCHGOPE0 " _
     & " where CHGOPEOPE in ('SWP' , 'TER')" _
     & " order by CHGOPEOPE , CHGOPEDOS "
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF

    If rsSab("CHGOPEENG") <= wAMJMin_IBM And rsSab("CHGOPEDT1") >= wAMJMin_IBM Then
         mXls1_Row = mXls1_Row + 1
         
         X = rsSab("CHGOPESER") & " " & rsSab("CHGOPESSE") & " " & rsSab("CHGOPEOPE") & " " & rsSab("CHGOPENAT") & "   " & Format(rsSab("CHGOPEDOS"), "### ###")
        Call lstErr_ChangeLastItem(lstErr, cmdContext, X): DoEvents
        
         'wsExcel.Cells(mXls1_Row, 1) = mId$(rsSab("YSTOPCI"), 1, 5)
         wsExcel.Cells(mXls1_Row, 3) = X
         wsExcel.Cells(mXls1_Row, 8) = dateImp10(rsSab("CHGOPEENG") + 19000000)
         wsExcel.Cells(mXls1_Row, 9) = dateImp10(rsSab("CHGOPEDT1") + 19000000)
         
        Call cmdPrint_Excel_ZCLIENA0(rsSab("CHGOPECON"), "     ", Trim(rsSab("CHGOPEOPE") & rsSab("CHGOPENAT")))
        wsExcel.Cells(mXls1_Row, 1) = "933000"
        
         wDev = rsSab("CHGOPEDE1")
         If rsSab("CHGOPEDE1") = "EUR" Then
            wDev = rsSab("CHGOPEDE2")
            wMTD = rsSab("CHGOPEMO2")
            wMTE = rsSab("CHGOPEMO1")
            wsExcel.Cells(mXls1_Row, 11) = rsSab("CHGOPECO2")
         Else
            wDev = rsSab("CHGOPEDE1")
            wMTD = rsSab("CHGOPEMO1")
            wMTE = rsSab("CHGOPEMO2")
            wsExcel.Cells(mXls1_Row, 11) = rsSab("CHGOPECO4")
            If rsSab("CHGOPESEN") = "V" Then wMTD = -wMTD
         End If
         
         wsExcel.Cells(mXls1_Row, 6) = wMTD
         wsExcel.Cells(mXls1_Row, 7) = wMTE
         
         wsExcel.Cells(mXls1_Row, 5) = wDev
         wsExcel.Cells(mXls1_Row, 12) = wMTD
        
        If wDev <> "EUR" Then
            Call sqlYBIATAB0_Read("PDC", wDev, wAMJMin, X)
            If IsNumeric(mId$(X, 9, 15)) Then
                dblCours = CDbl(mId$(X, 9, 15) / 1000000000)
                If dblCours <> 0 Then curX = Round(wMTD / dblCours, 2)
            Else
                curX = 0
            End If
        End If
         wsExcel.Cells(mXls1_Row, 13) = curX
         
    End If
    
'____________________________________________________________________________________

    rsSab.MoveNext
Loop




Exit Sub
'======================================================================================================

Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
Exit_sub:


End Sub

Public Sub cmdPrint_Excel_ZCLIENA0(lCLIENACLI As String, lPCI As String, lOPENAT As String)
Dim xSQL As String, X As String, K As Long, K2 As Long, xECO As String
On Error GoTo Error_Handler

'===================================================================================

     wsExcel.Cells(mXls1_Row, 4) = lCLIENACLI
    xSQL = "select * from " & paramIBM_Library_SAB & ".ZCLIENA0 " _
         & " where CLIENACLI = '" & lCLIENACLI & "'"
    Set rsSabX = cnsab.Execute(xSQL)
    
    If Not rsSabX.EOF Then
        wsExcel.Cells(mXls1_Row, 16) = Trim(rsSabX("CLIENARA1")) & Trim(rsSabX("CLIENARA2"))
        wsExcel.Cells(mXls1_Row, 2) = rsSabX("CLIENACAT") & " " & rsSabX("CLIENAECO") & " - " & rsSabX("CLIENARSD")
        
        wsExcel.Cells(mXls1_Row, 1) = lPCI & " " & PCI_Fiscal(Trim(rsSabX("CLIENARSD")))
        
        xECO = ""
        X = rsSabX("CLIENACAT")
        For K = 1 To arrCLIENACAT_Nb
            If X = arrCLIENACAT_Code(K) Then
                xECO = arrCLIENACAT_Lib(K)
            End If
        Next K
         X = rsSabX("CLIENAECO")
        For K = 1 To arrCLIENAECO_Nb
            If X = arrCLIENAECO_Code(K) Then
                xECO = xECO & arrCLIENAECO_Lib(K)
            End If
        Next K
        wsExcel.Cells(mXls1_Row, 15) = xECO
    End If
    
    
    For K = 1 To arrNature_Nb
        If lOPENAT = arrNature_Code(K) Then
            wsExcel.Cells(mXls1_Row, 14) = arrNature_Lib(K)
        End If
    Next K
'____________________________________________________________________________________

Exit Sub
'======================================================================================================

Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
Exit_sub:


End Sub

Public Sub cmdPrint_Excel_Init(lTitle As String)
Dim X As String, K As Long
On Error GoTo Error_Handler

'===================================================================================


With wsExcel.Cells
    .Borders.Weight = xlWide
    .Borders.Color = RGB(255, 255, 153)
    .Borders(xlInsideHorizontal).Weight = xlThin ' xlMedium
    .Borders(xlInsideHorizontal).Color = RGB(128, 128, 255)
    .Borders(xlInsideVertical).Weight = xlThin
    .Borders(xlInsideVertical).Color = RGB(128, 128, 255)
    .HorizontalAlignment = Excel.xlHAlignLeft
    .WrapText = False ' True
    .Font.Size = 9
    .Font.Name = "Calibri"
    .RowHeight = 17
End With

wsExcel.PageSetup.Orientation = vbPRORLandscape
wsExcel.PageSetup.HeaderMargin = 5
wsExcel.PageSetup.Zoom = 70

wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14" & lTitle & " , arrêté au " & dateImp10(wAMJMin) _
                                 & vbCr

wsExcel.PageSetup.CenterHorizontally = True


wsExcel.PageSetup.PrintTitleRows = "$A1:$K1"
wsExcel.PageSetup.LeftFooter = "&F - &B&A"
wsExcel.PageSetup.RightFooter = "&P / &N"

mXls1_Row = 1
mXls1_Cols = 16

wsExcel.Columns(1).ColumnWidth = 8: wsExcel.Cells(1, 1) = "PCI "
wsExcel.Columns(2).ColumnWidth = 13: wsExcel.Cells(1, 2) = "Cat client Pays"
wsExcel.Columns(3).ColumnWidth = 21: wsExcel.Cells(1, 3) = "Identifiant"
wsExcel.Columns(4).ColumnWidth = 8: wsExcel.Cells(1, 4) = "Racine"
wsExcel.Columns(5).ColumnWidth = 6: wsExcel.Cells(1, 5) = "Devise"
wsExcel.Columns(6).ColumnWidth = 18: wsExcel.Cells(1, 6) = "Montant d'origine"
wsExcel.Columns(6).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
wsExcel.Columns(6).HorizontalAlignment = Excel.xlHAlignRight
wsExcel.Columns(7).ColumnWidth = 18: wsExcel.Cells(1, 7) = "CV EUR d'origine"
wsExcel.Columns(7).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
wsExcel.Columns(7).HorizontalAlignment = Excel.xlHAlignRight
wsExcel.Columns(8).ColumnWidth = 11: wsExcel.Cells(1, 8) = "Date Début"
wsExcel.Columns(9).ColumnWidth = 11: wsExcel.Cells(1, 9) = "Date Fin"
wsExcel.Columns(10).ColumnWidth = 15: wsExcel.Cells(1, 10) = "type de taux"
wsExcel.Columns(11).ColumnWidth = 12: wsExcel.Cells(1, 11) = "Taux"
wsExcel.Columns(11).NumberFormat = "####0.000 00"
wsExcel.Columns(11).HorizontalAlignment = Excel.xlHAlignRight
wsExcel.Columns(12).ColumnWidth = 18: wsExcel.Cells(1, 12) = "Encours au " & dateImp10_S(wAMJMin)
wsExcel.Columns(12).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
wsExcel.Columns(12).HorizontalAlignment = Excel.xlHAlignRight
wsExcel.Columns(13).ColumnWidth = 18: wsExcel.Cells(1, 13) = "Encours EUR"
wsExcel.Columns(13).NumberFormat = "[Blue]### ### ### ##0.00;[Red]-### ### ### ##0.00"
wsExcel.Columns(13).HorizontalAlignment = Excel.xlHAlignRight
wsExcel.Columns(14).ColumnWidth = 24: wsExcel.Cells(1, 14) = "Nature produit"
wsExcel.Columns(15).ColumnWidth = 50: wsExcel.Cells(1, 15) = "Catégorie client"
wsExcel.Columns(16).ColumnWidth = 50: wsExcel.Cells(1, 16) = "Intitulé client"
wsExcel.Columns(16).Interior.Color = mColor_G0
wsExcel.Columns(4).Interior.Color = mColor_G0
wsExcel.Columns(3).Interior.Color = mColor_Y1
wsExcel.Columns(14).Interior.Color = mColor_Y1
wsExcel.Columns(2).Interior.Color = mColor_B0
wsExcel.Columns(15).Interior.Color = mColor_B0

For K = 1 To mXls1_Cols
    wsExcel.Cells(1, K).Interior.Color = mColor_GB
    wsExcel.Cells(1, K).Font.Color = vbWhite
Next


Exit Sub
'======================================================================================================

Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
Exit_sub:


End Sub


Public Function PCI_Fiscal(lPays As String) As String
Static K As Integer
If arrPays_Nb = 0 Then
    Call rsZBASTAB0_Pays(arrPays(), arrPays_Nb)
    K = 0: arrPays(0).Id = "?"
End If
'___________________________________________________________________________
If lPays = arrPays(K).Id Then
    PCI_Fiscal = arrPays(K).Fiscal
Else
    PCI_Fiscal = ""
    For K = 1 To arrPays_Nb
        If lPays = arrPays(K).Id Then PCI_Fiscal = arrPays(K).Fiscal: Exit For
    Next K
    If K > arrPays_Nb Then K = 1
End If
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




Public Sub fgdetail_Sort()
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



'___________________________________________________________________________
xSQL = "select count(*) from " & paramIBM_Library_SAB & ".ZCRETAB0 where CREtabeta = 1 and CRETABNUM = 9"
Set rsSab = cnsab.Execute(xSQL)
K = rsSab(0)
xSQL = "select count(*) from " & paramIBM_Library_SAB & ".ZBASTAB0 where bastabeta = 1 and BASTABNUM = 58"
Set rsSab = cnsab.Execute(xSQL)
K = K + rsSab(0) + 1
ReDim arrNature_Code(K), arrNature_Lib(K)
arrNature_Nb = 0
xSQL = "select * from " & paramIBM_Library_SAB & ".ZBASTAB0 where BASTABNUM = 58 order by BASTABARG"
Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    arrNature_Nb = arrNature_Nb + 1
    arrNature_Code(arrNature_Nb) = Trim(rsSab("BASTABARG"))
    arrNature_Lib(arrNature_Nb) = Trim(rsSab("BASTABDON"))
    rsSab.MoveNext
Loop
xSQL = "select * from " & paramIBM_Library_SAB & ".ZCRETAB0 where CREtabeta = 1 and CRETABNUM = 9 order by CRETABARG"
Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    arrNature_Nb = arrNature_Nb + 1
    arrNature_Code(arrNature_Nb) = "CRE" & Trim(rsSab("CRETABARG"))
    arrNature_Lib(arrNature_Nb) = Trim(mId$(rsSab("CRETABDON"), 1, 30))
    rsSab.MoveNext
Loop

'___________________________________________________________________________
xSQL = "select count(*) from " & paramIBM_Library_SAB & ".ZBASTAB0 where bastabeta = 1 and BASTABNUM = 8"
Set rsSab = cnsab.Execute(xSQL)
K = rsSab(0) + 1
ReDim arrCLIENACAT_Code(K), arrCLIENACAT_Lib(K)
arrCLIENACAT_Nb = 0
xSQL = "select * from " & paramIBM_Library_SAB & ".ZBASTAB0 where BASTABNUM = 8 order by BASTABARG"
Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    arrCLIENACAT_Nb = arrCLIENACAT_Nb + 1
    arrCLIENACAT_Code(arrCLIENACAT_Nb) = mId$(rsSab("BASTABARG"), 4, 3)
    arrCLIENACAT_Lib(arrCLIENACAT_Nb) = rsSab("BASTABLO2") & mId$(rsSab("BASTABDON"), 1, 15)
    rsSab.MoveNext
Loop
'___________________________________________________________________________
xSQL = "select count(*) from " & paramIBM_Library_SAB & ".ZBASTAB0 where bastabeta = 1 and BASTABNUM = 1"
Set rsSab = cnsab.Execute(xSQL)
K = rsSab(0) + 1
ReDim arrCLIENAECO_Code(K), arrCLIENAECO_Lib(K)
arrCLIENAECO_Nb = 0
xSQL = "select * from " & paramIBM_Library_SAB & ".ZBASTAB0 where BASTABNUM = 1 order by BASTABARG"
Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    arrCLIENAECO_Nb = arrCLIENAECO_Nb + 1
    arrCLIENAECO_Code(arrCLIENAECO_Nb) = mId$(rsSab("BASTABARG"), 4, 3)
    arrCLIENAECO_Lib(arrCLIENAECO_Nb) = rsSab("BASTABLO2") & mId$(rsSab("BASTABDON"), 1, 17)
    rsSab.MoveNext
Loop



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
fgSelect.Visible = False
fgSelect_Reset

fgSelect.Rows = 1
fgSelect.FormatString = "<CIB        |<Dénomination                                                                                |< Adresse                                                                                                          |<CP Ville                                         "
                 
fgSelect.Row = 0

Do While Not rsSab.EOF

    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1
    fgSelect_Display_Line
    
    rsSab.MoveNext

Loop

fgSelect.Visible = True

If fgSelect.Rows = 2 Then
    fgSelect.Col = 0
    Call fgDetail_Display(Trim(fgSelect.Text))
End If

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgSelect.Rows - 1): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Public Sub fgSelect_Display_Line()
Dim X As String

On Error Resume Next
'x = rsSab("ZFICBDF0")
'fgSelect.Col = 0: fgSelect.Text = mId$(x, 2, 5)
'fgSelect.Col = 1: fgSelect.Text = mId$(x, 13, 40)
'fgSelect.Col = 3: fgSelect.Text = mId$(x, 186, 32)
'fgSelect.Col = 2: fgSelect.Text = Trim(mId$(x, 90, 32 * 3))
End Sub

Public Sub Msg_Rcv(Msg As String)
'---------------------------------------------------------
Dim wFct As String

mWindowState = Me.WindowState
If Me.WindowState <> frmUsr_Windowstate Then Me.WindowState = frmUsr_Windowstate

wFct = UCase$(Trim(mId$(Msg, 1, 12)))
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
If Not txtRTF.Visible Then
    'cmdPrint_Display
Else
End If



End Sub

Private Sub fgDetail_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim wX As String, xUUMID As String
On Error Resume Next


If y <= fgDetail.RowHeightMin Then
    fgDetail.Visible = False
    Select Case fgDetail.Col
        Case 0: fgDetail_Sort1 = 0: fgDetail_Sort2 = 3: fgdetail_Sort
        Case 1:  fgDetail_Sort1 = 1: fgDetail_Sort2 = 3: fgdetail_Sort
        Case 2: fgDetail_Sort1 = 2: fgDetail_Sort2 = 2: fgdetail_Sort
        Case 3: fgDetail_Sort1 = 3: fgDetail_Sort2 = 3: fgdetail_Sort
        Case 4: fgDetail_Sort1 = 4: fgDetail_Sort2 = 4: fgdetail_Sort
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
            Case "2"
                fgSelect.Col = 0: wX = Trim(fgSelect.Text)
                Call fgDetail_Display(wX)
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
        cmdSelect_SQL_K = Trim(mId$(cboSelect_SQL, 1, K - 1))
    Else
        cmdSelect_SQL_K = "???"
    End If
    
    fraSelect_Options.Visible = False
    
    Select Case cmdSelect_SQL_K
        Case "1": cmdSelect_Ok.Visible = True: fraSelect_Options.Visible = True
    End Select

End If
End Sub


Private Sub cmdSelect_SQL_1()
Dim V, X As String
Dim xSQL As String, xWhere As String
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_1"
    
'xSQL = "select * from " & paramIBM_Library_SAB & ".ZFICBDF0 " & xWhere & " order by substring(ZFICBDF0 , 1 , 6 )"

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
Call lstErr_Clear(lstErr, cmdContext, "> BIA_VB_Habilitations_cmdSelect_Ok ........"): DoEvents

If fgSelect.Visible Then cmdSelect_Clear

Select Case cmdSelect_SQL_K
    Case "1": cmdPrint_Excel
    'Case "1": JPL_Cours
End Select
    
    
Call lstErr_AddItem(lstErr, cmdContext, "< BIA_VB_Habilitations_cmdSelect_Ok"): DoEvents
lstErr.Height = 480
Me.Enabled = True: Me.MousePointer = 0
If cmdSelect_Ok.Visible Then cmdSelect_Ok.SetFocus
cmdSelect_Ok.BackColor = fgSelect.BackColorFixed
End Sub




Public Function ZBASTAB0_Cours(lDEV As String, lAMJ As String) As Double
Dim wA1 As Integer, wA2 As Integer, wA3 As Integer, wA4 As Integer
Dim xAMJ As String, xSQL As String
Dim wAMJ_IBM As Long
Dim rsSab As New ADODB.Recordset

On Error GoTo Error_Handler
wAMJ_IBM = (Val(lAMJ) - 19000000)
ZBASTAB0_Cours = rsZBASTAB0_Cours37(lDEV, wAMJ_IBM)
If ZBASTAB0_Cours = 0 Then
    Select Case lDEV
        Case "CHF"
            Select Case lAMJ
                Case "20021001": ZBASTAB0_Cours = 1.4574
                Case "20021230": ZBASTAB0_Cours = 1.4548
            End Select
            
        Case "USD"
            Select Case lAMJ
            
                Case "20020915": ZBASTAB0_Cours = 0.981
                Case "20021101": ZBASTAB0_Cours = 0.9974
            End Select
    End Select

    If ZBASTAB0_Cours = 0 Then Call MsgBox(lDEV & lAMJ, vbCritical, "ZBASTAB0_Cours")
End If

Exit Function


'___________________________________________________________________
ZBASTAB0_Cours = 0
Call convX2P_IBMAMJ(xAMJ, wA1, wA2, wA3, wA4)
   
xSQL = "select * from " & paramIBM_Library_SAB & ".ZBASTAB0" _
     & " where BASTABETA = " & currentZMNURUT0.MNURUTETB _
     & " and   BASTABNUM = 37" _
     & " and substring(bastabarg , 1 , 3) = '" & lDEV & "'" _
                     & " and ascii(substring(bastablo1 , 1 , 1)) = " & wA1 _
                     & " and ascii(substring(bastablo1 , 2 , 1)) = " & wA2 _
                     & " and ascii(substring(bastablo1 , 3 , 1)) = " & wA3 _
                     & " and ascii(substring(bastablo1 , 4 , 1)) = " & wA4 _
     & " order by BASTABARG"


Set rsSab = cnsab.Execute(xSQL)
If rsSab.EOF Then
    xSQL = "select * from " & paramIBM_Library_SAB & ".ZBASTAB0" _
         & " where BASTABETA = " & currentZMNURUT0.MNURUTETB _
         & " and   BASTABNUM = 37" _
         & " and substring(bastabarg , 1 , 3) = '" & lDEV & "'" _
                         & " and ascii(substring(bastablo1 , 1 , 1)) = " & wA1 _
                         & " and ascii(substring(bastablo1 , 2 , 1)) = " & wA2 _
                         & " and ascii(substring(bastablo1 , 3 , 1)) = " & wA3 _
         & " order by BASTABARG"
    
    
    Set rsSab = cnsab.Execute(xSQL)
End If


Do While Not rsSab.EOF
    X = rsSab("BASTABARG")
    xAMJ = 19000000 + convX2P(mId$(X, 4, 4))
    If xAMJ = lAMJ Then
        If lDEV = mId$(X, 1, 3) Then
           ZBASTAB0_Cours = CDbl(convX2P(mId$(rsSab("BASTABDON"), 1, 8))) / 1000000000
           Exit Do
        End If
    End If
    rsSab.MoveNext

Loop

' PB récupération des cours
If ZBASTAB0_Cours = 0 Then

    Select Case lDEV
        Case "CHF"
            Select Case lAMJ
                Case "20030227": ZBASTAB0_Cours = 1.463
                Case "20031218": ZBASTAB0_Cours = 1.5539
                Case "20030930": ZBASTAB0_Cours = 1.5404
                Case "20021001": ZBASTAB0_Cours = 1.4574
                Case "20021230": ZBASTAB0_Cours = 1.4548
            End Select
            
        Case "USD"
            Select Case lAMJ
                Case "20070307": ZBASTAB0_Cours = 1.3135
                Case "20030207": ZBASTAB0_Cours = 1.0789
                Case "20040329": ZBASTAB0_Cours = 1.2118
                Case "20060131": ZBASTAB0_Cours = 1.2118
                Case "20030210": ZBASTAB0_Cours = 1.0808
            
                Case "20071025": ZBASTAB0_Cours = 1.4309
                Case "20040316": ZBASTAB0_Cours = 1.235
                Case "20060127": ZBASTAB0_Cours = 1.2172
            
                Case "20020915": ZBASTAB0_Cours = 0.981
                Case "20021101": ZBASTAB0_Cours = 0.9974
            End Select
    End Select
End If

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
    
Exit_sub:

End Function
Public Function JPL_Cours()
Dim wA1 As Integer, wA2 As Integer, wA3 As Integer, wA4 As Integer
Dim xAMJ As String, xSQL As String
Dim rsSab As New ADODB.Recordset
Dim K As Integer

Dim T_AMJ(5000) As String, T_Cours(5000) As Double, Nb As Integer
On Error GoTo Error_Handler

'GoTo X2

Debug.Print "37", rsZBASTAB0_Cours37("USD", 1030204)
    xAMJ = dateIBM("20030207")
    Call convX2P_IBMAMJ(xAMJ, wA1, wA2, wA3, wA4)
       
    xSQL = "select * from " & paramIBM_Library_SAB & ".ZBASTAB0" _
         & " where BASTABETA = " & currentZMNURUT0.MNURUTETB _
         & " and   BASTABNUM = 37" _
         & " and substring(bastabarg , 1 , 3) = 'USD'" _
                         & " and substring(bastabARG , 4, 4) = x'1030204F'" _
         & " order by BASTABARG"

    Set rsSab = cnsab.Execute(xSQL)
    
Do While Not rsSab.EOF
    X = rsSab("BASTABARG")
    xAMJ = 19000000 + convX2P(mId$(X, 4, 4))
    Debug.Print xAMJ, CDbl(convX2P(mId$(rsSab("BASTABDON"), 1, 8))) / 1000000000
    rsSab.MoveNext

Loop
GoTo Exit_sub
'_______________________________________________________
X2:

xSQL = "select * from " & paramIBM_Library_SAB & ".ZBASTAB0" _
     & " where BASTABETA = " & currentZMNURUT0.MNURUTETB _
     & " and   BASTABNUM = 37" _
     & " and substring(bastabarg , 1 , 3) = 'USD'" _
     & " order by BASTABARG"

Set rsSab = cnsab.Execute(xSQL)


Do While Not rsSab.EOF
    X = rsSab("BASTABARG")
    xAMJ = 19000000 + convX2P(mId$(X, 4, 4))
    Nb = Nb + 1
    T_AMJ(Nb) = xAMJ
    T_Cours(Nb) = CDbl(convX2P(mId$(rsSab("BASTABDON"), 1, 8))) / 1000000000
    'Debug.Print xAMJ, CDbl(convX2P(mId$(rsSab("BASTABDON"), 1, 8))) / 1000000000
    rsSab.MoveNext

Loop

For K = 1 To Nb
    xAMJ = dateIBM(T_AMJ(K))
    'Call convX2P_IBMAMJ(xAMJ, wA1, wA2, wA3, wA4)
       
    xSQL = "select * from " & paramIBM_Library_SAB & ".ZBASTAB0" _
         & " where BASTABETA = " & currentZMNURUT0.MNURUTETB _
         & " and   BASTABNUM = 37" _
         & " and substring(bastabarg , 1 , 3) = 'USD'" _
                         & " and substring(bastabARG , 4, 4) = x'" & xAMJ & "F'" _
         & " order by BASTABARG"

    Set rsSab = cnsab.Execute(xSQL)
    
    If Not rsSab.EOF Then
        X = rsSab("BASTABARG")
        xAMJ = 19000000 + convX2P(mId$(X, 4, 4))
        Nb = Nb + 1
        If xAMJ <> T_AMJ(K) Then
            Debug.Print K, T_AMJ(K), xAMJ
        End If
    Else
    
            Debug.Print K, T_AMJ(K), "??"
    
    End If
    
Next K

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
    
Exit_sub:

End Function


