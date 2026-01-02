VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmSIDE_DB 
   AutoRedraw      =   -1  'True
   Caption         =   "SAA : historique des messages SWIFT"
   ClientHeight    =   12180
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   17730
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SIDE_DB.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   12180
   ScaleWidth      =   17730
   Begin VB.TextBox txtFg 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3210
      Left            =   420
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   8
      Text            =   "SIDE_DB.frx":030A
      Top             =   6045
      Visible         =   0   'False
      Width           =   6555
   End
   Begin MSFlexGridLib.MSFlexGrid fgSAA_Detail 
      Height          =   2160
      Left            =   1815
      TabIndex        =   3
      Top             =   4005
      Visible         =   0   'False
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   3810
      _Version        =   393216
      Cols            =   4
      FixedCols       =   2
      RowHeightMin    =   300
      BackColor       =   16316664
      ForeColor       =   8192
      BackColorFixed  =   13693183
      ForeColorFixed  =   0
      BackColorSel    =   13693183
      BackColorBkg    =   16316664
      WordWrap        =   -1  'True
      GridLines       =   3
      AllowUserResizing=   3
      FormatString    =   $"SIDE_DB.frx":0312
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame fraSwift 
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
      Height          =   11475
      Left            =   60
      TabIndex        =   4
      Top             =   540
      Width           =   17505
      Begin RichTextLib.RichTextBox txtRTF 
         Height          =   3945
         Left            =   9150
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   5085
         Visible         =   0   'False
         Width           =   6105
         _ExtentX        =   10769
         _ExtentY        =   6959
         _Version        =   393217
         BackColor       =   15790320
         HideSelection   =   0   'False
         ScrollBars      =   3
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"SIDE_DB.frx":03C8
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
      Begin MSFlexGridLib.MSFlexGrid fgSwift 
         Height          =   10830
         Left            =   90
         TabIndex        =   5
         Top             =   510
         Width           =   7110
         _ExtentX        =   12541
         _ExtentY        =   19103
         _Version        =   393216
         FixedCols       =   0
         RowHeightMin    =   400
         BackColor       =   16777215
         ForeColor       =   12582912
         BackColorFixed  =   12648447
         ForeColorFixed  =   0
         BackColorBkg    =   16777215
         GridColor       =   12632064
         GridColorFixed  =   12632064
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         FormatString    =   $"SIDE_DB.frx":0448
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid fgSAA_Histo 
         Height          =   8970
         Left            =   7260
         TabIndex        =   6
         Top             =   2430
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   15822
         _Version        =   393216
         Cols            =   13
         FixedCols       =   0
         RowHeightMin    =   350
         BackColor       =   16777215
         ForeColor       =   16711680
         BackColorFixed  =   8421376
         ForeColorFixed  =   16777215
         BackColorBkg    =   16777215
         GridColor       =   10526720
         GridColorFixed  =   10526720
         AllowUserResizing=   3
         FormatString    =   $"SIDE_DB.frx":04D5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid fgSAB_Histo 
         Height          =   1875
         Left            =   7230
         TabIndex        =   10
         Top             =   525
         Width           =   10200
         _ExtentX        =   17992
         _ExtentY        =   3307
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         RowHeightMin    =   350
         BackColor       =   16777215
         ForeColor       =   16711680
         BackColorFixed  =   12640511
         ForeColorFixed  =   0
         BackColorBkg    =   16777215
         GridColor       =   10526720
         GridColorFixed  =   10526720
         AllowUserResizing=   3
         FormatString    =   $"SIDE_DB.frx":05B0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label libSWIFT_SWISABSWID 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   195
         TabIndex        =   7
         Top             =   195
         Width           =   17085
      End
   End
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   480
      Left            =   10170
      TabIndex        =   2
      Top             =   60
      Width           =   6900
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
      Left            =   17145
      Picture         =   "SIDE_DB.frx":0670
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   15
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
   Begin VB.Menu mnuZSWIBIC0 
      Caption         =   "BIC"
      Visible         =   0   'False
      Begin VB.Menu mnuSWIBICBIC 
         Caption         =   "BIC"
      End
   End
End
Attribute VB_Name = "frmSIDE_DB"
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
Dim YGOSDOS0_Aut As typeAuthorization
Dim blnAuto As Boolean, blnError As Boolean
Dim cmdSelect_SQL_K As String

Dim rsSabX As New ADODB.Recordset

Dim oldYSWISAB0 As typeYSWISAB0
Dim fgSwift_FormatString As String
Dim cnSIDE_DB As New ADODB.Connection, rsSIDE_DB As New ADODB.Recordset
Dim xrText As typerText

Dim fgSAA_Histo_FormatString As String, fgSAA_Histo_K As Integer
Dim fgSAA_Histo_RowDisplay As Integer, fgSAA_Histo_RowClick As Integer, fgSAA_Histo_ColClick As Integer
Dim fgSAA_Histo_ColorClick As Long, fgSAA_Histo_ColorDisplay As Long
Dim fgSAA_Histo_Sort1 As Integer, fgSAA_Histo_Sort2 As Integer
Dim fgSAA_Histo_SortAD As Integer, fgSAA_Histo_Sort1_Old As Integer
Dim fgSAA_Histo_arrIndex As Integer
Dim blnfgSAA_Histo_DisplayLine As Boolean

Dim fgSAB_Histo_FormatString As String, fgSAB_Histo_K As Integer
Dim fgSAB_Histo_RowDisplay As Integer, fgSAB_Histo_RowClick As Integer, fgSAB_Histo_ColClick As Integer
Dim fgSAB_Histo_ColorClick As Long, fgSAB_Histo_ColorDisplay As Long
Dim fgSAB_Histo_Sort1 As Integer, fgSAB_Histo_Sort2 As Integer
Dim fgSAB_Histo_SortAD As Integer, fgSAB_Histo_Sort1_Old As Integer
Dim fgSAB_Histo_arrIndex As Integer
Dim blnfgSAB_Histo_DisplayLine As Boolean

Dim fgSAA_Detail_FormatString As String, fgSAA_Detail_K As Integer
Dim fgSAA_Detail_RowDisplay As Integer, fgSAA_Detail_RowClick As Integer, fgSAA_Detail_ColClick As Integer
Dim fgSAA_Detail_ColorClick As Long, fgSAA_Detail_ColorDisplay As Long
Dim fgSAA_Detail_Sort1 As Integer, fgSAA_Detail_Sort2 As Integer
Dim fgSAA_Detail_SortAD As Integer, fgSAA_Detail_Sort1_Old As Integer
Dim fgSAA_Detail_arrIndex As Integer
Dim blnfgSAA_Detail_DisplayLine As Boolean
Dim fgSAA_Detail_BackColorFixed As Long

Dim arrMesg() As String, arrMesg_Nb As Integer
Dim arrAppe() As String, arrAppe_Nb As Integer
Dim arrIntv() As String, arrIntv_Nb As Integer
Dim arrInst() As String, arrInst_Nb As Integer
Dim inst_num As Long, Inst_date_time As String, Inst_seq_nbr As Long
Dim HeightOfLine As Long, LinesOfText As Long

Dim txtRTF_prtForeColor_Header As Long
Dim Mesg_aid As Long, mesg_s_umidl As Long, mesg_s_umidh As Long

Public Sub fgSwift_Display(lSWISABSWID As Long, lMesg_aid As Long, lmesg_s_umidl As Long, lmesg_s_umidh As Long)
Dim wColor As Long, wColorFixed As Long
Dim X As String, xWhere As String, xOPE As String
Dim xSql As String
Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean
Dim wAmj As String
Dim xUUMID As String
On Error Resume Next

If Not frmSIDE_DB.Visible Then frmSIDE_DB.Visible = True
frmSIDE_DB.Show

txtRTF.Visible = False

X = Trim(frmSIDE_DB.Caption)
AppActivate X
On Error GoTo Error_Handler

fraSwift.Visible = False: fgSAA_Detail.Visible = False: txtFg.Visible = False
'fgswift_Reset
libSWIFT_SWISABSWID = "BIA Id : " & lSWISABSWID
fgSwift.Rows = 1
fgSwift.FormatString = fgSwift_FormatString
fgSwift.Row = 0
fgSwift.RowHeight(0) = 700

fgSAB_Histo.Clear: fgSAB_Histo.FormatString = fgSAB_Histo_FormatString

currentAction = "fgswift_Display"

'----------------------------------------------------------------
Mesg_aid = lMesg_aid
mesg_s_umidl = lmesg_s_umidl
mesg_s_umidh = lmesg_s_umidh

blnOk = False
If lSWISABSWID > 0 Then
    xSql = "select * from " & paramIBM_Library_SABSPE & ".YSWISAB0 where SWISABSWID = " & lSWISABSWID
    Set rsSab = cnsab.Execute(xSql)
    
    If Not rsSab.EOF Then
        blnOk = True
        Call rsYSWISAB0_GetBuffer(rsSab, oldYSWISAB0)
        Mesg_aid = oldYSWISAB0.SWISABWID1
        mesg_s_umidl = oldYSWISAB0.SWISABWIDL
        mesg_s_umidh = oldYSWISAB0.SWISABWIDH

    End If
End If
'----------------------------------------------------------------
If Not blnOk Then
    Call rsYSWISAB0_Init(oldYSWISAB0)
    libSWIFT_SWISABSWID = " !!! inconnu dans YSWISAB0 et SAA !!!!!!!!!!!!!!!"
    xSql = "select * from rMesg " _
        & "where Aid = " & Mesg_aid _
        & " and Mesg_s_umidl = " & mesg_s_umidl _
        & " and Mesg_s_umidh  =  " & mesg_s_umidh
   Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)
   
    If Not rsSIDE_DB.EOF Then
        If Not IsNull(rsSIDE_DB("mesg_type")) Then
            oldYSWISAB0.SWISABWMTK = rsSIDE_DB("mesg_type")
        Else
            oldYSWISAB0.SWISABWMTK = "???"
        End If
         xUUMID = rsSIDE_DB("mesg_uumid")
         If Mid$(xUUMID, 1, 1) = "I" Then
             oldYSWISAB0.SWISABWES = "S"
         Else
             oldYSWISAB0.SWISABWES = "E"
         End If
        oldYSWISAB0.SWISABWBIC = Mid$(xUUMID, 2, 11)
        Call dateJma10_Amj(Mid$(rsSIDE_DB("mesg_crea_date_time"), 1, 10), X)
        oldYSWISAB0.SWISABWAMJ = Val(X)
        X = Mid$(rsSIDE_DB("mesg_crea_date_time"), 12, 8)
        oldYSWISAB0.SWISABWHMS = Val(Mid$(X, 1, 2) & Mid$(X, 4, 2) & Mid$(X, 7, 2))

    End If
End If
'--------------------------------------------------------------
If blnOk And oldYSWISAB0.SWISABZSWI Then
    fgSAB_Histo_Display
End If

'--------------------------------------------------------------
    
If oldYSWISAB0.SWISABWES = "E" Then
    wColor = RGB(190, 240, 255)
    wColorFixed = vbBlue
    X = "reçu de "
    libSWIFT_SWISABSWID = "Réception MT " & oldYSWISAB0.SWISABWMTK & " de "
    fgSAA_Detail_BackColorFixed = mColor_B9
Else
    wColor = RGB(220, 255, 220)
    wColorFixed = RGB(0, 64, 0)
    X = "émis vers "
    libSWIFT_SWISABSWID = "Emission MT " & oldYSWISAB0.SWISABWMTK & " vers "
    fgSAA_Detail_BackColorFixed = mColor_G9
End If
libSWIFT_SWISABSWID = libSWIFT_SWISABSWID & oldYSWISAB0.SWISABWBIC & "   le " & dateImp10(oldYSWISAB0.SWISABWAMJ) & "   " & timeImp8(oldYSWISAB0.SWISABWHMS)
If oldYSWISAB0.SWISABOPEN <> 0 Then
    libSWIFT_SWISABSWID = libSWIFT_SWISABSWID & "     Dossier : " & oldYSWISAB0.SWISABOPEC & " " & Format(oldYSWISAB0.SWISABOPEN, "######")
End If

libSWIFT_SWISABSWID.ForeColor = wColorFixed
libSWIFT_SWISABSWID.BackColor = wColor
'fgSwift.Col = 0: fgSwift.CellBackColor = wColor: fgSwift.Text = oldYSWISAB0.SWISABWMTK
'fgSwift.CellFontBold = True
'fgSwift.Col = 1: fgSwift.CellBackColor = wColor: fgSwift.Text = Trim(oldYSWISAB0.SWISABOPEC) & " " & Format(oldYSWISAB0.SWISABOPEN, "######")
'fgSwift.CellFontBold = True
fgSwift.Col = 0: fgSwift.Text = oldYSWISAB0.SWISABWMTK
fgSwift.CellFontBold = True: fgSwift.CellBackColor = wColor
fgSwift.ForeColorFixed = wColorFixed
fgSwift.Col = 1: fgSwift.Text = X & oldYSWISAB0.SWISABWBIC & " le " & dateImp10(oldYSWISAB0.SWISABWAMJ) & " " & timeImp8(oldYSWISAB0.SWISABWHMS) _
                                  & vbCrLf & ZSWIBIC0_Select(oldYSWISAB0.SWISABWBIC)
fgSwift.CellFontBold = True: fgSwift.CellBackColor = wColor
fgSwift.ForeColorFixed = wColorFixed

xSql = "select * from rtextField " _
    & "where Aid = " & Mesg_aid _
    & " and text_s_umidl = " & mesg_s_umidl _
    & " and text_s_umidh  =  " & mesg_s_umidh _
    & " order by field_cnt"

Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)
If Not rsSIDE_DB.EOF Then
    Do While Not rsSIDE_DB.EOF
    
        fgSwift.Rows = fgSwift.Rows + 1
        fgSwift.Row = fgSwift.Rows - 1
    
        fgSwift_DisplayLine fgSwift.Row, wColor, wColorFixed
    
        rsSIDE_DB.MoveNext
    
    Loop
Else
    xSql = "select * from rtext " _
        & "where Aid = " & Mesg_aid _
        & " and text_s_umidl = " & mesg_s_umidl _
        & " and text_s_umidh  =  " & mesg_s_umidh
    Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)
    If Not rsSIDE_DB.EOF Then
        Call srvrText_GetBuffer_ODBC(rsSIDE_DB, xrText)
        fgSwift_DisplayLine_rText fgSwift.Row, wColor, wColorFixed
    End If
End If
fraSwift.Visible = True

fgSAA_Histo_Display


'___________________________________________________________________________

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Public Function ZSWIBIC0_Select(lMsg As String) As String
Dim xSql As String
xSql = "select * from " & paramIBM_Library_SAB & ".ZSWIBIC0 where SWIBICBIC like '" & Trim(lMsg) & "%' order by SWIBICBIC"
Set rsSabX = cnsab.Execute(xSql)

If Not rsSabX.EOF Then
    ZSWIBIC0_Select = Trim(rsSabX("SWIBICIN1")) & "  " & Trim(rsSabX("SWIBICVIL")) & "  " & Trim(rsSabX("SWIBICCOM"))
Else
    ZSWIBIC0_Select = ""
End If

End Function


Private Sub fgSAA_Detail_Display_rAppe()
Dim xSql As String
On Error GoTo Error_Handler

Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
fgSAA_Detail.Visible = False: txtFg.Visible = False
fgSAA_Detail_Reset

fgSAA_Detail.Rows = 1
fgSAA_Detail.FormatString = fgSAA_Detail_FormatString
fgSAA_Detail.Row = 0

currentAction = "fgSAA_Detail_Display"

If arrAppe_Nb = 0 Then
    xSql = "SELECT    count(*) From sysobjects, syscolumns " _
          & " WHERE  ( sysobjects.id = syscolumns.id) And  (sysobjects.xtype = 'U') " _
          & " AND sysobjects.name LIKE 'rappe'"
    Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)
    arrAppe_Nb = rsSIDE_DB(0)
    ReDim arrAppe(arrAppe_Nb + 1)
    I = 0
    
    xSql = "SELECT    syscolumns.name From sysobjects, syscolumns " _
          & " WHERE  ( sysobjects.id = syscolumns.id) And  (sysobjects.xtype = 'U') " _
          & " AND sysobjects.name LIKE 'rappe' ORDER BY syscolumns.colorder"
    Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)
    Do While Not rsSIDE_DB.EOF
            arrAppe(I) = rsSIDE_DB(0)
            I = I + 1
        rsSIDE_DB.MoveNext
    Loop
End If


X = "select * from rappe " _
    & "where Aid = " & Mesg_aid _
    & " and appe_s_umidl = " & mesg_s_umidl _
    & " and appe_s_umidh  =  " & mesg_s_umidh _
    & " and appe_inst_num  =  " & inst_num _
    & " and appe_date_time  =  '" & Inst_date_time & "'" _
    & " and appe_seq_nbr  =  " & Inst_seq_nbr
    
    
Set rsSIDE_DB = cnSIDE_DB.Execute(X)
If Not rsSIDE_DB.EOF Then
    For I = 0 To arrAppe_Nb - 1
        fgSAA_Detail.Rows = fgSAA_Detail.Rows + 1
        fgSAA_Detail.Row = fgSAA_Detail.Rows - 1
        
        fgSAA_Detail.Col = 0: fgSAA_Detail.Text = I
        fgSAA_Detail.Col = 1: fgSAA_Detail.Text = arrAppe(I)
        fgSAA_Detail.Col = 2
        V = rsSIDE_DB(I)
        If Not IsNull(V) Then
            fgSAA_Detail.Text = Trim(V)
            If Len(fgSAA_Detail.Text) > 65 Then fgSAA_Detail_Display_Text
            
        End If
        
    Next I
End If


fgSAA_Detail.Visible = True

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgSAA_Detail.Rows - 1): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub


Public Sub fgSAA_Detail_Reset()
fgSAA_Detail.Clear
fgSAA_Detail_Sort1 = 0: fgSAA_Detail_Sort2 = 0
fgSAA_Detail_Sort1_Old = -1
fgSAA_Detail_RowDisplay = 0: fgSAA_Detail_RowClick = 0
fgSAA_Detail_arrIndex = fgSAA_Detail.Cols - 1
blnfgSAA_Detail_DisplayLine = False
fgSAA_Detail_SortAD = 6
fgSAA_Detail.LeftCol = fgSAA_Detail.FixedCols

End Sub


Public Sub fgSAA_Histo_Sort()
If fgSAA_Histo.Rows > 1 Then
    fgSAA_Histo.Row = 1
    fgSAA_Histo.RowSel = fgSAA_Histo.Rows - 1
    
    If fgSAA_Histo_Sort1_Old = fgSAA_Histo_Sort1 Then
        If fgSAA_Histo_SortAD = 5 Then
            fgSAA_Histo_SortAD = 6
        Else
            fgSAA_Histo_SortAD = 5
        End If
    Else
        fgSAA_Histo_SortAD = 5
    End If
    fgSAA_Histo_Sort1_Old = fgSAA_Histo_Sort1
    
    fgSAA_Histo.Col = fgSAA_Histo_Sort1
    fgSAA_Histo.ColSel = fgSAA_Histo_Sort2
    fgSAA_Histo.Sort = fgSAA_Histo_SortAD
End If

End Sub

Private Sub fgSAA_Detail_Display_rIntv()
Dim xSql As String
On Error GoTo Error_Handler

Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
fgSAA_Detail.Visible = False: txtFg.Visible = False
fgSAA_Detail_Reset

fgSAA_Detail.Rows = 1
fgSAA_Detail.FormatString = fgSAA_Detail_FormatString
fgSAA_Detail.Row = 0

currentAction = "fgSAA_Detail_Display"

If arrIntv_Nb = 0 Then
    xSql = "SELECT    count(*) From sysobjects, syscolumns " _
          & " WHERE  ( sysobjects.id = syscolumns.id) And  (sysobjects.xtype = 'U') " _
          & " AND sysobjects.name LIKE 'rIntv'"
    Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)
    arrIntv_Nb = rsSIDE_DB(0)
    ReDim arrIntv(arrIntv_Nb + 1)
    I = 0
    
    xSql = "SELECT    syscolumns.name From sysobjects, syscolumns " _
          & " WHERE  ( sysobjects.id = syscolumns.id) And  (sysobjects.xtype = 'U') " _
          & " AND sysobjects.name LIKE 'rIntv' ORDER BY syscolumns.colorder"
    Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)
    Do While Not rsSIDE_DB.EOF
            arrIntv(I) = rsSIDE_DB(0)
            I = I + 1
        rsSIDE_DB.MoveNext
    Loop
End If


X = "select * from rIntv " _
    & "where Aid = " & Mesg_aid _
    & " and Intv_s_umidl = " & mesg_s_umidl _
    & " and Intv_s_umidh  =  " & mesg_s_umidh _
    & " and Intv_inst_num  =  " & inst_num _
    & " and Intv_date_time  =  '" & Inst_date_time & "'" _
    & " and Intv_seq_nbr  =  " & Inst_seq_nbr
    
    
Set rsSIDE_DB = cnSIDE_DB.Execute(X)
If Not rsSIDE_DB.EOF Then
    For I = 0 To arrIntv_Nb - 1
        fgSAA_Detail.Rows = fgSAA_Detail.Rows + 1
        fgSAA_Detail.Row = fgSAA_Detail.Rows - 1
        
        fgSAA_Detail.Col = 0: fgSAA_Detail.Text = I
        fgSAA_Detail.Col = 1: fgSAA_Detail.Text = arrIntv(I)
        fgSAA_Detail.Col = 2
        V = rsSIDE_DB(I)
        If Not IsNull(V) Then
            fgSAA_Detail.Text = Trim(V)
            If Len(fgSAA_Detail.Text) > 65 Then fgSAA_Detail_Display_Text
            
        End If
            
    Next I
End If

fgSAA_Detail.Visible = True

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgSAA_Detail.Rows - 1): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub


Public Sub fgSAA_Histo_SortX(lK As Integer)
Dim I As Integer, X As String, wIndex As Long

For I = 1 To fgSAA_Histo.Rows - 1
    fgSAA_Histo.Row = I
    fgSAA_Histo.Col = lK
    Select Case lK
'        Case 3: fgSAA_Histo.Col = 3: X = Format$(Val(fgSAA_Histo.Text), "000000000000000.00")

    End Select
    fgSAA_Histo.Col = fgSAA_Histo_arrIndex - 1
    fgSAA_Histo.Text = X
Next I

fgSAA_Histo_Sort1 = fgSAA_Histo_arrIndex - 1: fgSAA_Histo_Sort2 = fgSAA_Histo_arrIndex - 1
fgSAA_Histo_Sort
End Sub





Private Sub fgSAA_Detail_Display_rInst()
Dim xSql As String
On Error GoTo Error_Handler

Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
fgSAA_Detail.Visible = False: txtFg.Visible = False
fgSAA_Detail_Reset

fgSAA_Detail.Rows = 1
fgSAA_Detail.FormatString = fgSAA_Detail_FormatString
fgSAA_Detail.Row = 0

currentAction = "fgSAA_Detail_Display"

If arrInst_Nb = 0 Then
    xSql = "SELECT    count(*) From sysobjects, syscolumns " _
          & " WHERE  ( sysobjects.id = syscolumns.id) And  (sysobjects.xtype = 'U') " _
          & " AND sysobjects.name LIKE 'rInst'"
    Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)
    arrInst_Nb = rsSIDE_DB(0)
    ReDim arrInst(arrInst_Nb + 1)
    I = 0
    
    xSql = "SELECT    syscolumns.name From sysobjects, syscolumns " _
          & " WHERE  ( sysobjects.id = syscolumns.id) And  (sysobjects.xtype = 'U') " _
          & " AND sysobjects.name LIKE 'rInst' ORDER BY syscolumns.colorder"
    Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)
    Do While Not rsSIDE_DB.EOF
            arrInst(I) = rsSIDE_DB(0)
            I = I + 1
        rsSIDE_DB.MoveNext
    Loop
End If


X = "select * from rInst " _
    & "where Aid = " & Mesg_aid _
    & " and inst_s_umidl = " & mesg_s_umidl _
    & " and inst_s_umidh  =  " & mesg_s_umidh _
    & " and inst_num  =  " & inst_num
    
    
Set rsSIDE_DB = cnSIDE_DB.Execute(X)
If Not rsSIDE_DB.EOF Then
    For I = 0 To arrInst_Nb - 1
        fgSAA_Detail.Rows = fgSAA_Detail.Rows + 1
        fgSAA_Detail.Row = fgSAA_Detail.Rows - 1
        
        fgSAA_Detail.Col = 0: fgSAA_Detail.Text = I
        fgSAA_Detail.Col = 1: fgSAA_Detail.Text = arrInst(I)
        fgSAA_Detail.Col = 2
        V = rsSIDE_DB(I)
        If Not IsNull(V) Then
            fgSAA_Detail.Text = Trim(V)
            If Len(fgSAA_Detail.Text) > 65 Then fgSAA_Detail_Display_Text
            
        End If
        
    Next I
End If


fgSAA_Detail.Visible = True

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgSAA_Detail.Rows - 1): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Public Sub fgSAA_Histo_Reset()
fgSAA_Histo.Clear
fgSAA_Histo_Sort1 = 0: fgSAA_Histo_Sort2 = 0
fgSAA_Histo_Sort1_Old = -1
fgSAA_Histo_RowDisplay = 0: fgSAA_Histo_RowClick = 0
fgSAA_Histo_arrIndex = fgSAA_Histo.Cols - 1
blnfgSAA_Histo_DisplayLine = False
fgSAA_Histo_SortAD = 6
fgSAA_Histo.LeftCol = fgSAA_Histo.FixedCols

End Sub



Public Sub fgSAB_Histo_Reset()
fgSAB_Histo.Clear
fgSAB_Histo_Sort1 = 0: fgSAB_Histo_Sort2 = 0
fgSAB_Histo_Sort1_Old = -1
fgSAB_Histo_RowDisplay = 0: fgSAB_Histo_RowClick = 0
fgSAB_Histo_arrIndex = fgSAB_Histo.Cols - 1
blnfgSAB_Histo_DisplayLine = False
fgSAB_Histo_SortAD = 6
fgSAB_Histo.LeftCol = fgSAB_Histo.FixedCols

End Sub
Private Sub fgSAA_Detail_Display_rMesg()
Dim xSql As String
On Error GoTo Error_Handler

Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
fgSAA_Detail.Visible = False: txtFg.Visible = False
fgSAA_Detail_Reset

fgSAA_Detail.Rows = 1
fgSAA_Detail.FormatString = fgSAA_Detail_FormatString
fgSAA_Detail.Row = 0

currentAction = "fgSAA_Detail_Display"

If arrMesg_Nb = 0 Then
    xSql = "SELECT    count(*) From sysobjects, syscolumns " _
          & " WHERE  ( sysobjects.id = syscolumns.id) And  (sysobjects.xtype = 'U') " _
          & " AND sysobjects.name LIKE 'rMesg'"
    Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)
    arrMesg_Nb = rsSIDE_DB(0)
    ReDim arrMesg(arrMesg_Nb + 1)
    I = 0
    
    xSql = "SELECT    syscolumns.name From sysobjects, syscolumns " _
          & " WHERE  ( sysobjects.id = syscolumns.id) And  (sysobjects.xtype = 'U') " _
          & " AND sysobjects.name LIKE 'rMesg' ORDER BY syscolumns.colorder"
    Set rsSIDE_DB = cnSIDE_DB.Execute(xSql)
    Do While Not rsSIDE_DB.EOF
            arrMesg(I) = rsSIDE_DB(0)
            I = I + 1
        rsSIDE_DB.MoveNext
    Loop
End If



X = "select * from rMesg " _
    & "where Aid = " & Mesg_aid _
    & " and mesg_s_umidl = " & mesg_s_umidl _
    & " and mesg_s_umidh  =  " & mesg_s_umidh
    
    
Set rsSIDE_DB = cnSIDE_DB.Execute(X)
If Not rsSIDE_DB.EOF Then
    For I = 0 To arrMesg_Nb - 1
        fgSAA_Detail.Rows = fgSAA_Detail.Rows + 1
        fgSAA_Detail.Row = fgSAA_Detail.Rows - 1
        
        fgSAA_Detail.Col = 0: fgSAA_Detail.Text = I
        fgSAA_Detail.Col = 1: fgSAA_Detail.Text = arrMesg(I)
        fgSAA_Detail.Col = 2
        V = rsSIDE_DB(I)
        If Not IsNull(V) Then
            fgSAA_Detail.Text = Trim(V)
            If Len(fgSAA_Detail.Text) > 65 Then fgSAA_Detail_Display_Text
            
        End If
        
    Next I
End If



fgSAA_Detail.Visible = True

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgSAA_Detail.Rows - 1): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Public Sub fgSAA_Histo_Display_rMesg()
Dim V, X As String, wColor As Long

X = "select * from rMesg " _
    & "where Aid = " & Mesg_aid _
    & " and mesg_s_umidl = " & mesg_s_umidl _
    & " and mesg_s_umidh  =  " & mesg_s_umidh
Set rsSIDE_DB = cnSIDE_DB.Execute(X)
If Trim(rsSIDE_DB("mesg_status")) = "COMPLETED" Then
    wColor = libSWIFT_SWISABSWID.BackColor 'mColor_G1
Else
    wColor = mColor_W0
    libSWIFT_SWISABSWID.BackColor = mColor_W0
End If
'libSWIFT_SWISABSWID.BackColor = wColor
Do While Not rsSIDE_DB.EOF

    fgSAA_Histo.Rows = fgSAA_Histo.Rows + 1
    fgSAA_Histo.Row = fgSAA_Histo.Rows - 1
    fgSAA_Histo.Col = 0: V = rsSIDE_DB(66)
    If Not IsNull(V) Then fgSAA_Histo.Text = V
    fgSAA_Histo.CellBackColor = wColor
    fgSAA_Histo.Col = 1: V = rsSIDE_DB(16): If Not IsNull(V) Then fgSAA_Histo.Text = V
    fgSAA_Histo.CellBackColor = wColor
    fgSAA_Histo.Col = 2: V = rsSIDE_DB(11): If Not IsNull(V) Then X = Replace(V, "COMPLETED", "terminé"): fgSAA_Histo.Text = X
    fgSAA_Histo.CellBackColor = wColor
    fgSAA_Histo.Col = 3: V = rsSIDE_DB(70): If Not IsNull(V) Then fgSAA_Histo.Text = V
    fgSAA_Histo.CellBackColor = wColor
    
    fgSAA_Histo.Col = 4: fgSAA_Histo.Text = "M"
    fgSAA_Histo.CellBackColor = wColor
    rsSIDE_DB.MoveNext

Loop


End Sub


Public Sub fgSAA_Histo_Display_rIntv()
Dim V, X As String, blnOFCS As Boolean, blnOFCS_Update As Boolean, K As Integer

X = "select * from rIntv " _
    & "where Aid = " & Mesg_aid _
    & " and Intv_s_umidl = " & mesg_s_umidl _
    & " and Intv_s_umidh  =  " & mesg_s_umidh _
    & " order by intv_inst_num,intv_date_time,intv_seq_nbr"
    
Set rsSIDE_DB = cnSIDE_DB.Execute(X)
    
Do While Not rsSIDE_DB.EOF

    fgSAA_Histo.Rows = fgSAA_Histo.Rows + 1
    fgSAA_Histo.Row = fgSAA_Histo.Rows - 1
    
    blnOFCS = False: blnOFCS_Update = False
    fgSAA_Histo.Col = 1: V = rsSIDE_DB(4): If Not IsNull(V) Then fgSAA_Histo.Text = V
    fgSAA_Histo.Col = 2: V = rsSIDE_DB(11)
    If Not IsNull(V) Then
        Select Case Trim(V)
            Case "OFCS_Detect": fgSAA_Histo.Text = V: blnOFCS = True
            Case "OFCS_Update": fgSAA_Histo.Text = V: blnOFCS_Update = True
            'Case "mpm": V = rsSIDE_DB(10): If Not IsNull(V) Then fgSAA_Histo.Text = V
            Case "mpa", "mpm": V = rsSIDE_DB(9): If Not IsNull(V) Then fgSAA_Histo.Text = V: fgSAA_Histo.CellBackColor = RGB(255, 255, 128)
            Case "AI_from_APPLI", "AI_to_APPLI"
            Case "_SI_to_SWIFT": fgSAA_Histo.Text = "=>Swift"
            Case "_SI_from_SWIFT": fgSAA_Histo.Text = "<=Swift"
            Case Else: fgSAA_Histo.Text = V
        End Select
    End If
    fgSAA_Histo.Col = 3: V = rsSIDE_DB(17)
    If Not IsNull(V) Then
        fgSAA_Histo.Text = V

        If blnOFCS And InStr(V, "Detection report") > 0 Then
            fgSAA_Histo.Col = 2
            If InStr(V, "No_Violation") > 0 Then
                fgSAA_Histo.CellBackColor = mColor_G1
            Else
                fgSAA_Histo.CellBackColor = RGB(255, 190, 96)

            End If
        End If
        If blnOFCS_Update And InStr(V, "Routed from") > 0 Then
            fgSAA_Histo.Col = 2
            If InStr(V, "[_MP_verification]") > 0 Then '"_SI_to_SWIFT"
                fgSAA_Histo.CellBackColor = mColor_G2
            Else
                If InStr(V, "[_SI_to_SWIFT]") > 0 Then
                    fgSAA_Histo.CellBackColor = mColor_G2
                Else
                    fgSAA_Histo.CellBackColor = mColor_W1
                End If

            End If
        End If
    End If
    
    
    fgSAA_Histo.Col = 4: fgSAA_Histo.Text = "V"
    fgSAA_Histo.Col = 5: V = rsSIDE_DB(3): If Not IsNull(V) Then fgSAA_Histo.Text = Format(V, "00000")
    fgSAA_Histo.Col = 6: V = rsSIDE_DB(4)
    If Not IsNull(V) Then
        fgSAA_Histo.Text = V
        'fgSAA_Histo.Text = Mid$(V, 7, 4) & Mid$(V, 4, 2) & Mid$(V, 1, 2) & Mid$(V, 12, 8)
    End If

    fgSAA_Histo.Col = 7: V = rsSIDE_DB(5): If Not IsNull(V) Then fgSAA_Histo.Text = Format(V, "0000000000")
    rsSIDE_DB.MoveNext

Loop


End Sub

Public Sub fgSAA_Histo_Display_rAppe()
Dim V, X As String, I As Integer
Dim blnSwift As Boolean, x6 As String

X = "select * from rAppe " _
    & "where Aid = " & Mesg_aid _
    & " and Appe_s_umidl = " & mesg_s_umidl _
    & " and Appe_s_umidh  =  " & mesg_s_umidh _
    & " order by Appe_inst_num,Appe_date_time,Appe_seq_nbr"
    
Set rsSIDE_DB = cnSIDE_DB.Execute(X)
    
Do While Not rsSIDE_DB.EOF

    fgSAA_Histo.Rows = fgSAA_Histo.Rows + 1
    fgSAA_Histo.Row = fgSAA_Histo.Rows - 1
    blnSwift = False
    fgSAA_Histo.Col = 0: fgSAA_Histo.Text = "+"
    fgSAA_Histo.Col = 1: V = rsSIDE_DB(4): If Not IsNull(V) Then fgSAA_Histo.Text = V
    
    V = rsSIDE_DB(6): If Not IsNull(V) Then x6 = Trim(V)
    fgSAA_Histo.Col = 2
    If x6 = "APPLI" Then
         V = rsSIDE_DB(8)
         If Not IsNull(V) Then
            fgSAA_Histo.Text = V
            If Trim(V) = "FileSabOutput" Then
                    For I = 0 To 10: fgSAA_Histo.Col = I: fgSAA_Histo.CellBackColor = RGB(128, 255, 128): Next I
            End If
        End If

   Else
        V = rsSIDE_DB(13)
        If Not IsNull(V) Then
            Select Case Trim(V)
                Case "AI_to_APPLI":
                Case "_SI_to_SWIFT": fgSAA_Histo.Text = "=>Swift": blnSwift = True
                Case "_SI_from_SWIFT": fgSAA_Histo.Text = "<=Swift": blnSwift = True
                Case Else: fgSAA_Histo.Text = V
            End Select
        End If
    End If
    
    fgSAA_Histo.Col = 3: V = rsSIDE_DB(35)
    If Not IsNull(V) Then
        fgSAA_Histo.Text = V
        Select Case Trim(V)
            Case "DLV_NACKED":
                    For I = 0 To 10: fgSAA_Histo.Col = I: fgSAA_Histo.CellBackColor = mColor_W1: Next I
            Case "DLV_ACKED":
                    If blnSwift Then
                        For I = 0 To 10: fgSAA_Histo.Col = I: fgSAA_Histo.CellBackColor = RGB(128, 255, 128): Next I
                    Else
                        fgSAA_Histo.Text = "Ok"
                    End If
                    
        End Select
    End If
    
    fgSAA_Histo.Col = 4: fgSAA_Histo.Text = "A"
    fgSAA_Histo.Col = 5: V = rsSIDE_DB(3): If Not IsNull(V) Then fgSAA_Histo.Text = Format(V, "00000")
    fgSAA_Histo.Col = 6: V = rsSIDE_DB(4)
    If Not IsNull(V) Then
        fgSAA_Histo.Text = V
        'fgSAA_Histo.Text = Mid$(V, 7, 4) & Mid$(V, 4, 2) & Mid$(V, 1, 2) & Mid$(V, 12, 8)
    End If
    fgSAA_Histo.Col = 7: V = rsSIDE_DB(5): If Not IsNull(V) Then fgSAA_Histo.Text = Format(V, "0000000000")
    rsSIDE_DB.MoveNext

Loop


End Sub


Public Sub fgSAA_Detail_Display_Text()
         
txtFg.Height = fgSAA_Detail.RowHeightMin
txtFg.Width = fgSAA_Detail.ColWidth(2)
txtFg = Trim(fgSAA_Detail.Text)
HeightOfLine = fgSAA_Detail.RowHeightMin / 1.3 - 20
LinesOfText = SendMessage(txtFg.hwnd, EM_GETLINECOUNT, 0&, 0&) + 1
If fgSAA_Detail.RowHeight(fgSAA_Detail.Row) < (LinesOfText * HeightOfLine) Then
   fgSAA_Detail.RowHeight(fgSAA_Detail.Row) = LinesOfText * HeightOfLine
   If fgSAA_Detail.RowHeight(fgSAA_Detail.Row) > 10000 Then
        fgSAA_Detail.RowHeight(fgSAA_Detail.Row) = 10000
        fgSAA_Detail.CellBackColor = mColor_Y1
    End If
End If
            
End Sub

Public Sub fgSAA_Histo_Display_rInst()
Dim V, X As String, X10 As String

X = "select * from rInst " _
    & "where Aid = " & Mesg_aid _
    & " and inst_s_umidl = " & mesg_s_umidl _
    & " and inst_s_umidh  =  " & mesg_s_umidh
Set rsSIDE_DB = cnSIDE_DB.Execute(X)
    
Do While Not rsSIDE_DB.EOF

    fgSAA_Histo.Rows = fgSAA_Histo.Rows + 1
    fgSAA_Histo.Row = fgSAA_Histo.Rows - 1
    fgSAA_Histo.Col = 0: V = rsSIDE_DB(4)
    If Not IsNull(V) Then
        Select Case Trim(V)
            Case "INST_TYPE_NOTIFICATION": fgSAA_Histo.Text = "notif"
            Case "INST_TYPE_ORIGINAL": fgSAA_Histo.Text = "original"
            Case "INST_TYPE_COPY": fgSAA_Histo.Text = "copy"
            Case Else: fgSAA_Histo.Text = V
        End Select
    End If
    fgSAA_Histo.CellBackColor = RGB(255, 255, 210)
    fgSAA_Histo.Col = 1: V = rsSIDE_DB(22): If Not IsNull(V) Then fgSAA_Histo.Text = V
    fgSAA_Histo.CellBackColor = RGB(255, 255, 210)
    
    fgSAA_Histo.Col = 2: V = rsSIDE_DB(12)
    fgSAA_Histo.CellBackColor = RGB(255, 255, 210)
    If Not IsNull(V) Then
        Select Case Trim(V)
            Case "AI_to_APPLI":
            Case "_SI_to_SWIFT": fgSAA_Histo.Text = "=>Swift"
            Case "_SI_from_SWIFT": fgSAA_Histo.Text = "<=Swift"
            Case Else: fgSAA_Histo.Text = V
                        If Trim(rsSIDE_DB(6)) = "COMPLETED" Then fgSAA_Histo.Text = "terminé"
        End Select
    End If
    V = rsSIDE_DB(10)
    If Not IsNull(V) Then
        X10 = Trim(V)
    Else
        X10 = ""
    End If
    fgSAA_Histo.Col = 3: V = rsSIDE_DB(15): If Not IsNull(V) Then X = Replace(V, "R_SUCCESS", "Ok"): fgSAA_Histo.Text = X10 & " - " & X
    fgSAA_Histo.CellBackColor = RGB(255, 255, 210)
    fgSAA_Histo.Col = 4: fgSAA_Histo.Text = "I"
    fgSAA_Histo.CellBackColor = RGB(255, 255, 210)
    fgSAA_Histo.Col = 5: V = rsSIDE_DB(3): If Not IsNull(V) Then fgSAA_Histo.Text = Format(V, "00000")
    fgSAA_Histo.CellBackColor = RGB(255, 255, 210)
    rsSIDE_DB.MoveNext

Loop


End Sub

'______________________________________________________________________
Private Sub fgSAA_Histo_Display()

Dim K As Long

On Error GoTo Error_Handler
fgSAA_Histo.Visible = False
fgSAA_Histo_Reset

fgSAA_Histo.Rows = 1
fgSAA_Histo.FormatString = fgSAA_Histo_FormatString
fgSAA_Histo.Row = 0
For K = 1 To fgSAA_Histo.Cols
    fgSAA_Histo.BackColorFixed = fgSAA_Detail_BackColorFixed
Next K
currentAction = "fgSAA_Histo_Display"

fgSAA_Histo_Display_rMesg
fgSAA_Histo_Display_rInst
fgSAA_Histo_Display_rIntv
fgSAA_Histo_Display_rAppe


fgSAA_Histo_Sort1 = 5: fgSAA_Histo_Sort2 = 7: fgSAA_Histo_Sort

fgSAA_Histo.Visible = True

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgSAA_Histo.Rows - 1): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub
Public Sub fgSAA_Histo_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
fgSAA_Histo.Visible = False
mRow = fgSAA_Histo.Row

If lRow > 0 And lRow < fgSAA_Histo.Rows Then
    fgSAA_Histo.Row = lRow
    For I = 1 To fgSAA_Histo.FixedCols Step -1
        fgSAA_Histo.Col = I: fgSAA_Histo.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgSAA_Histo.Row = mRow
    If fgSAA_Histo.Row > 0 Then
        lRow = fgSAA_Histo.Row
        lColor_Old = fgSAA_Histo.CellBackColor
        For I = 1 To fgSAA_Histo.FixedCols Step -1
          fgSAA_Histo.Col = I: fgSAA_Histo.CellBackColor = lColor
        Next I
    End If
End If
fgSAA_Histo.LeftCol = fgSAA_Histo.FixedCols
fgSAA_Histo.Visible = True
End Sub


Public Sub fgSwift_DisplayLine(lIndex As Long, lCellBackColor As Long, lColorFixed As Long)
Dim K As Integer, iAsc13 As Integer, iLen As Integer
Dim blnAsc13 As Boolean
Dim xValue As String
Dim V
On Error Resume Next
fgSwift.Col = 0: fgSwift.Text = rsSIDE_DB("field_code") & rsSIDE_DB("field_option")
fgSwift.CellBackColor = lCellBackColor
fgSwift.Col = 1
fgSwift.CellForeColor = lColorFixed

        Select Case rsSIDE_DB("field_code")
            Case "45", "46", "47", "77":
                V = rsSIDE_DB("value_memo")
                If IsNull(V) Then V = rsSIDE_DB("value")
            Case Else:
                    V = rsSIDE_DB("value")
        End Select
        If IsNull(V) Then
            xValue = ""
        Else
            xValue = V
        End If


 iLen = Len(xValue)
 K = 1
 Do
    iAsc13 = InStr(K, xValue, Asc13)
    If iAsc13 > 0 Then
        fgSwift.Text = Trim(Mid$(xValue, K, iAsc13 - K))
        fgSwift.CellForeColor = lColorFixed
        If Len(fgSwift.Text) > 45 Then fgSwift.RowHeight(fgSwift.Row) = 500
        K = iAsc13 + 2
        fgSwift.Rows = fgSwift.Rows + 1
        fgSwift.Row = fgSwift.Rows - 1
    End If
 Loop Until iAsc13 = 0

fgSwift.Text = Trim(Mid$(xValue, K, iLen - K + 1))
fgSwift.CellForeColor = lColorFixed
'fgSwift.Col = fgSwift.Cols - 1: fgSwift.Text = rsSIDE_DB("field_cnt")


End Sub

Public Sub fgSwift_DisplayLine_rText(lIndex As Long, lCellBackColor As Long, lColorFixed As Long)
Dim K As Integer, iAsc13 As Integer, iLen As Integer
Dim blnAsc13 As Boolean
Dim xValue As String, X As String, K2 As Integer
Dim blnField_79 As Boolean

blnField_79 = False
On Error Resume Next

xValue = xrText.text_data_block & Asc13
 iLen = Len(xValue)
If Mid$(xValue, 1, 3) = Asc13 & Asc10 & ":" Then
    K = 3
Else
    K = 1
End If
 Do
    iAsc13 = InStr(K, xValue, Asc13)
    If iAsc13 > 0 Then
        fgSwift.Rows = fgSwift.Rows + 1
        fgSwift.Row = fgSwift.Rows - 1
        X = Trim(Mid$(xValue, K, iAsc13 - K))
        fgSwift.Col = 1
        If Mid$(X, 1, 1) <> ":" Then
            fgSwift.Text = Trim(Mid$(xValue, K, iAsc13 - K))
            fgSwift.CellForeColor = lColorFixed
        Else
            'K2 = InStr(2, x, ":")
            If blnField_79 Then
                K2 = 0
            Else
                K2 = InStr(2, X, ":")
            End If
            If K2 > 0 Then
                fgSwift.Text = Trim(Mid$(X, K2 + 1, Len(X) - K2))
                fgSwift.CellForeColor = lColorFixed
                fgSwift.Col = 0: fgSwift.Text = Trim(Mid$(X, 2, K2 - 2))
                If Trim(Mid$(X, 2, K2 - 2)) = "79" Then blnField_79 = True

                fgSwift.CellBackColor = lCellBackColor
            Else
                fgSwift.Text = Trim(Mid$(xValue, K, iAsc13 - K))
                fgSwift.CellForeColor = lColorFixed
            End If
        End If
        
        K = iAsc13 + 2
    End If
 Loop Until iAsc13 = 0

'fgSwift.Text = Trim(Mid$(xValue, K, iLen - K + 1))
'fgSwift.CellForeColor = lColorFixed
'fgSwift.Col = fgSwift.Cols - 1: fgSwift.Text = rsSIDE_DB("field_cnt")


End Sub



Private Sub cmdContext_Click()
Unload Me
End Sub

Private Sub cmdPrint_Click()
If Not txtRTF.Visible Then
    cmdPrint_Display
Else
    prtOrientation = vbPRORLandscape
    prtEdition_Open
    prtTitleText = libSWIFT_SWISABSWID
    prtPgmName = "frmSIDE_DB"
    prtHeaderHeight = 300
    prtFormType = ""
    'prtForeColor_Header = txtRTF_prtForeColor_Header
    frmElpPrt.prtStdInit
    XPrt.CurrentY = prtMinY + prtHeaderHeight
    Call frmElpPrt.prtRTF(txtRTF.TextRTF)
    
    XPrt.DrawWidth = 5

    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)

    prtEdition_Close
End If



End Sub
Private Sub cmdPrint_Display()

Dim K As Integer, xRTF As String, X As String, XL As String
Dim kSelStart As Long
Dim arrColor_SelStart(1000) As Long, arrColor_SelLength(1000) As Long, arrColor(1000) As Long, arrColor_Nb As Long
Dim arrBold_SelStart(1000) As Long, arrBold_SelLength(1000) As Long, arrBold_Nb As Long
Dim arrUnderline_SelStart(1000) As Long, arrUnderline_SelLength(1000) As Long, arrUnderline_Nb As Long
Dim xTab As String, K1 As Integer, K2 As Integer, L1 As Integer, L2 As Integer, xLine_Len As Integer
Dim blnExit As Boolean, mLen As Integer, xLine As String
Dim mSelstart As Long, blnColor As Boolean, wColor As Long

xTab = vbCrLf & vbTab & vbTab & vbTab & vbTab & vbTab
X = "MT " & oldYSWISAB0.SWISABWMTK
'_____________________________________________________________
If oldYSWISAB0.SWISABWES = "E" Then
    X = "Réception MT " & oldYSWISAB0.SWISABWMTK & " de "
    arrColor(1) = vbBlue
Else
    X = "Emission MT " & oldYSWISAB0.SWISABWMTK & " vers "
    arrColor(1) = RGB(0, 128, 0)
End If
X = X & oldYSWISAB0.SWISABWBIC & "   le " & dateImp10(oldYSWISAB0.SWISABWAMJ) & "   " & timeImp8(oldYSWISAB0.SWISABWHMS)
txtRTF_prtForeColor_Header = arrColor(1)

arrBold_SelStart(1) = 0
arrBold_SelLength(1) = Len(X)
arrBold_Nb = 1
arrColor_Nb = 1
arrColor_SelStart(1) = 1
arrColor_SelLength(1) = Len(X)

xRTF = X & vbTab & vbTab & vbTab
'_____________________________________________________

If oldYSWISAB0.SWISABOPEN <> 0 Then
    X = "Dossier : " & oldYSWISAB0.SWISABSER & " " & oldYSWISAB0.SWISABSSE & " " & oldYSWISAB0.SWISABOPEC & " " & oldYSWISAB0.SWISABOPEN
    arrColor_Nb = arrColor_Nb + 1
    arrColor_SelStart(arrColor_Nb) = Len(xRTF)
    arrColor_SelLength(arrColor_Nb) = Len(X)
    arrColor(arrColor_Nb) = RGB(255, 0, 128)
Else
    X = ""
End If
xRTF = xRTF & X & vbCrLf & String$(150, "_") & vbCrLf
'=============================================================================

For K = 1 To fgSwift.Rows - 1
    fgSwift.Row = K
    fgSwift.Col = 0
    
    X = Trim(fgSwift.Text) & vbTab & ": "
    If Trim(fgSwift.Text) <> "" Then
        arrColor_Nb = arrColor_Nb + 1
        arrColor_SelStart(arrColor_Nb) = Len(xRTF)
        arrColor_SelLength(arrColor_Nb) = Len(X)
        arrColor(arrColor_Nb) = RGB(96, 96, 96)
    End If
    xRTF = xRTF & X

    fgSwift.Col = 1
    X = Trim(fgSwift)
    arrColor_Nb = arrColor_Nb + 1
    arrColor_SelStart(arrColor_Nb) = Len(xRTF)
    arrColor_SelLength(arrColor_Nb) = Len(X)
    arrColor(arrColor_Nb) = vbBlue
    xRTF = xRTF & X & vbCrLf
Next K

xRTF = xRTF & vbCrLf & vbCrLf
'=============================================================================
'On ne prend pas l'historique car trop volumineux DR 11/12/2018
GoTo PAS_DHISTORIQUE
For K = 1 To fgSAA_Histo.Rows - 1
    fgSAA_Histo.Row = K
    fgSAA_Histo.Col = 0
    
    If Trim(fgSAA_Histo.Text) = "" Then
        X = "   " & vbTab & ": "
    Else
        X = Trim(fgSAA_Histo.Text) & vbTab & ": "
    End If
    'arrBold_Nb = arrColor_Nb + 1
    'arrBold_SelStart(arrColor_Nb) = Len(xRTF)
    'arrBold_SelLength(arrColor_Nb) = Len(X)
    
    fgSAA_Histo.Col = 1
    X = X & Trim(fgSAA_Histo) & ":  "

    fgSAA_Histo.Col = 2
    X = X & Trim(fgSAA_Histo) & vbTab
    If Len(Trim(fgSAA_Histo)) < 5 Then X = X & vbTab
    
    fgSAA_Histo.Col = 3
       
    blnColor = False
    mSelstart = Len(xRTF)
    XL = Replace(Trim(fgSAA_Histo), vbCrLf, xTab & vbTab) & vbCrLf
    
    If InStr(XL, "Detection report") > 0 Then
        blnColor = True: wColor = vbMagenta
        If Len(XL) > 250 Then XL = Mid$(XL, 1, 250) & xTab & vbTab & "........................................." & vbCrLf
    End If
    If InStr(XL, "Modified data") > 0 Then blnColor = True: wColor = mColor_B9

    blnExit = False: K1 = 1: mLen = Len(XL)
    Do
        K2 = InStr(K1, XL, vbCrLf)
        
        If K2 > 0 Then
            xLine = Mid$(XL, K1, K2 - K1 + 2)
            xLine_Len = Len(xLine)
            For L1 = 1 To xLine_Len Step 120
                L2 = Len(xLine) - L1 + 1
                If L2 > 120 Then
                    X = X & Mid$(xLine, L1, 120) & xTab
                Else
                    X = X & Mid$(xLine, L1, L2)
                End If
            Next L1
            K1 = K2 + 3
        End If
        
        If K1 >= mLen Then blnExit = True
    Loop Until blnExit
    If blnColor Then
        arrColor_Nb = arrColor_Nb + 1
        arrColor_SelStart(arrColor_Nb) = mSelstart
        arrColor_SelLength(arrColor_Nb) = Len(XL)
        arrColor(arrColor_Nb) = wColor
    End If
    
    blnColor = False
    XL = ""
    Select Case fgSAA_Histo.CellBackColor
   
        Case mColor_G1, RGB(220, 255, 220): blnColor = True: wColor = RGB(0, 96, 0): XL = String$(150, "_") & vbCrLf
        Case RGB(255, 255, 210): blnColor = True: wColor = RGB(128, 64, 0): XL = String$(150, "_") & vbCrLf
        Case RGB(190, 240, 255): blnColor = True: wColor = vbBlue
        Case RGB(128, 255, 128): blnColor = True: wColor = vbBlue
        Case mColor_W0, mColor_W1: blnColor = True: wColor = RGB(255, 0, 128): XL = String$(150, "_") & vbCrLf
    End Select

     If blnColor Then
        arrColor_Nb = arrColor_Nb + 1
        arrColor_SelStart(arrColor_Nb) = mSelstart
        arrColor_SelLength(arrColor_Nb) = Len(XL & X)
        arrColor(arrColor_Nb) = wColor
    End If
        
    
    xRTF = xRTF & XL & X & vbCrLf
    
Next K
PAS_DHISTORIQUE:
'=============================================================================

txtRTF = xRTF
txtRTF.Font.Name = "Calibri"
txtRTF.Font.Size = 8
For K = 1 To arrColor_Nb
    txtRTF.SelStart = arrColor_SelStart(K)
    txtRTF.SelLength = arrColor_SelLength(K)
    txtRTF.SelColor = arrColor(K)
    'Debug.Print K, arrColor_SelStart(K), arrColor_SelLength(K), arrColor(K)
Next K

For K = 1 To arrBold_Nb
    txtRTF.SelStart = arrBold_SelStart(K)
    txtRTF.SelLength = arrBold_SelLength(K)
    txtRTF.SelBold = True
Next K

For K = 1 To arrUnderline_Nb
    txtRTF.SelStart = arrUnderline_SelStart(K)
    txtRTF.SelLength = arrUnderline_SelLength(K)
    txtRTF.SelUnderline = True
Next K

txtRTF.Locked = True
txtRTF.Visible = True

End Sub

Private Sub fgSAA_Detail_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim wX As String, K As Integer, K2 As Integer
Dim wX2 As String
On Error Resume Next


If y <= fgSAA_Detail.RowHeightMin Then
Else
    If fgSAA_Detail.Rows > 1 Then
        fgSAA_Detail.Col = 2: wX = Trim(fgSAA_Detail.Text)
        wX2 = ""
        For K = 1 To Len(wX)
            wX2 = wX2 & Mid$(wX, K, 1)
            If Mid$(wX, K, 1) = vbCr Then
                K2 = 0
            Else
                If K2 > 75 Then
                    wX2 = wX2 & vbCrLf: K2 = 0
                Else
                    K2 = K2 + 1
                End If
            End If
        Next K
        txtFg = wX2
        txtFg.Visible = True
    txtFg.Top = fraSwift.Top
    txtFg.Height = fraSwift.Height
    txtFg.Width = 9000
    txtFg.Left = fraSwift.Left + fraSwift.Width - txtFg.Width
        
   End If
End If
fgSAA_Detail.LeftCol = 0

End Sub


Private Sub fgSAA_Histo_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim wX As String, xUUMID As String
On Error Resume Next


If y <= fgSAA_Histo.RowHeightMin Then
Else
    If fgSAA_Histo.Rows > 1 Then
        Call fgSAA_Histo_Color(fgSAA_Histo_RowClick, MouseMoveUsr.BackColor, fgSAA_Histo_ColorClick)
        fgSAA_Histo.Col = 4: wX = Trim(fgSAA_Histo.Text)
        fgSAA_Histo.Col = 5: inst_num = Val(fgSAA_Histo.Text)
        fgSAA_Histo.Col = 6: Inst_date_time = Trim(fgSAA_Histo.Text)
        fgSAA_Histo.Col = 7: Inst_seq_nbr = Val(fgSAA_Histo.Text)
        Select Case wX
            Case "M": fgSAA_Detail_Display_rMesg
            Case "I": fgSAA_Detail_Display_rInst
            Case "V": fgSAA_Detail_Display_rIntv
            Case "A": fgSAA_Detail_Display_rAppe
        End Select
   End If
End If
fgSAA_Histo.LeftCol = 0


End Sub

Private Sub fgSwift_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim xField As String, xSql As String
If fgSwift.Rows > 1 Then
    If X <= 500 Then
    fgSwift.Col = 0
        xField = Trim(fgSwift.Text)
        Call arrMT_Type_Scan(oldYSWISAB0.SWISABWMTK)
        mnuSWIBICBIC.Caption = arrMT_Fields_Scan(xField)
        Me.PopupMenu mnuZSWIBIC0, vbPopupMenuLeftButton
    Else
        fgSwift.Col = 1
        If ZSWIBIC0_Select(fgSwift.Text) <> "" Then
            mnuSWIBICBIC.Caption = Trim(rsSabX("SWIBICIN1")) & "  " & Trim(rsSabX("SWIBICVIL")) & "  " & Trim(rsSabX("SWIBICCOM"))
            Me.PopupMenu mnuZSWIBIC0, vbPopupMenuLeftButton
        End If
    End If
End If
fgSwift.Col = 0

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

'---------------------------------------------------------
Public Sub arrTag_Set(ByVal B As Boolean)
'---------------------------------------------------------
For I = 0 To arrTagNb: arrTag(I) = B: Next I
End Sub


Public Sub cmdContext_Return()
        SendKeys "{TAB}"
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

If fgSAA_Detail.Visible Then
    fgSAA_Detail.Visible = False
    Exit Sub
End If


'If fraSwift.Visible Then
'    fraSwift.Visible = False
'    Exit Sub
'End If

Unload Me

End Sub

Private Sub Form_Load()

frmSIDE_DB_Show

Set XForm = Me
Me.Left = 19000 - Me.Width
KeyPreview = True

blnControl = False
cnSIDE_DB.Open paramODBC_DSN_SIDE_DB
'mWindowState = Me.WindowState
mHeight_0 = Me.Height: mWidth_0 = Me.Width: mHeight_2 = 0: mWidth_2 = 0: mWindowState = Me.WindowState
If Me.WindowState <> frmUsr_Windowstate Then Me.WindowState = frmUsr_Windowstate

fgSwift_FormatString = fgSwift.FormatString

fgSAA_Histo_FormatString = fgSAA_Histo.FormatString
fgSAB_Histo_FormatString = fgSAB_Histo.FormatString
fgSAA_Histo.Enabled = True
fgSAA_Histo.Visible = False

fgSAA_Detail_FormatString = fgSAA_Detail.FormatString
fgSAA_Detail.Enabled = True
fgSAA_Detail.Visible = False: txtFg.Visible = False
fgSAA_Detail.Top = fraSwift.Top
fgSAA_Detail.Height = fraSwift.Height
fgSAA_Detail.Width = 9000
fgSAA_Detail.Left = fraSwift.Left + 100
libSWIFT_SWISABSWID.BackColor = mColor_Y1
libSWIFT_SWISABSWID.ForeColor = RGB(128, 64, 0)
txtRTF.Top = 1005
txtRTF.Left = 240
txtRTF.Height = 8025
txtRTF.Width = 13020
End Sub


Private Sub Form_Resize()
If mWindowState <> Me.WindowState Then
    If Me.WindowState = 0 Or Me.WindowState = 2 Then
        Elp_Form_Resize Me, mWindowState, mHeight_0, mWidth_0, mHeight_2, mWidth_2
    End If
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
cnSIDE_DB.Close
Set cnSIDE_DB = Nothing

End Sub


Public Sub fgSAB_Histo_Display()

On Error GoTo Error_Handler
fgSAB_Histo.Visible = False
fgSAB_Histo_Reset

fgSAB_Histo.Rows = 1
fgSAB_Histo.FormatString = fgSAB_Histo_FormatString
fgSAB_Histo.Row = 0

currentAction = "fgSAB_Histo_Display"

If paramEnvironnement = constTest Then paramIBM_Library_SABJRN = "SAB073JRN"

Select Case oldYSWISAB0.SWISABWES
    Case "S": fgSAB_Histo_Display_S
    Case "E": fgSAB_Histo_Display_E
End Select
'________________________________________________________________________________________________
fgSAB_Histo.Visible = True

Call lstErr_AddItem(lstErr, cmdContext, "enregistrements : " & fgSAB_Histo.Row): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub
Public Sub fgSAB_Histo_Display_S()
Dim xSql As String, wColor As Long
Dim blnOk As Boolean, blnDisplay As Boolean
Dim mJOENTT As String, mJOUSER As String
Dim xProcess As String
On Error GoTo Error_Handler

xSql = "select * from " & paramIBM_Library_SABJRN & ".JSWIFTA0 W , " & paramIBM_Library_SABJRN & ".JRNENT0 J" _
     & " Where W.JORCV = J.JORCV And W.JOSEQN = J.JOSEQN" _
     & " and SWIFTANUM = " & oldYSWISAB0.SWISABZSWI _
     & " order by W.JORCV , W.JOSEQN"
'xSQL = "select * from " & paramIBM_Library_SABJRN & ".JSWIFTA0 W , " & paramIBM_Library_SABJRN & ".JRNENT0 J" _
'     & " Where SWIFTANUM = " & oldYSWISAB0.SWISABZSWI _
'     & " and W.JORCV = J.JORCV And W.JOSEQN = J.JOSEQN" _
'     & " order by W.JORCV , W.JOSEQN"
blnOk = False
Set rsSab = cnsab.Execute(xSql)
If Not rsSab.EOF Then
'________________________________________________________________________________________________
    Do While Not rsSab.EOF
        mJOENTT = rsSab("JOENTT")
        wColor = RGB(255, 255, 255)
        blnDisplay = False
        
        Select Case mJOENTT
            Case "PX": xProcess = "création": blnDisplay = True
            Case "UP": xProcess = "modification"
                       If mJOUSER <> rsSab("JOUSER") Then blnDisplay = True: wColor = mColor_Y1
                       If rsSab("SWIFTAVAL") = "O" Then blnDisplay = True: xProcess = "validation": wColor = mColor_G1
            Case "DL":
                If rsSab("SWIFTAVAL") <> "O" Then xProcess = "suppression": wColor = mColor_W0: blnDisplay = True
        End Select
        If blnDisplay Then
            mJOUSER = rsSab("JOUSER")
            fgSAB_Histo.Rows = fgSAB_Histo.Rows + 1
            fgSAB_Histo.Row = fgSAB_Histo.Rows - 1
            fgSAB_Histo.Col = 0: fgSAB_Histo.Text = "J": fgSAB_Histo.CellBackColor = wColor
            fgSAB_Histo.Col = 1: fgSAB_Histo.Text = Format(rsSab("JODATE"), "00/00/00") & "  " & Format(rsSab("JOTIME"), "00:00:00")
            fgSAB_Histo.CellBackColor = wColor
            fgSAB_Histo.Col = 2: fgSAB_Histo.Text = xProcess: fgSAB_Histo.CellBackColor = wColor
            fgSAB_Histo.Col = 3: fgSAB_Histo.Text = mJOUSER: fgSAB_Histo.CellBackColor = wColor
        End If
        rsSab.MoveNext
    Loop
'________________________________________________________________________________________________
Else
    xSql = "select * from " & paramIBM_Library_SAB & ".ZSWIHIA0" _
         & " Where SWIHIANUM = " & oldYSWISAB0.SWISABZSWI
    Set rsSab = cnsab.Execute(xSql)
    
    Do While Not rsSab.EOF
        
        If rsSab("SWIHIAVAL") = "O" Then
            xProcess = "validé": wColor = mColor_Y1
        Else
            xProcess = "???": wColor = mColor_W0
        End If
        fgSAB_Histo.Rows = fgSAB_Histo.Rows + 1
        fgSAB_Histo.Row = fgSAB_Histo.Rows - 1
        fgSAB_Histo.Col = 0: fgSAB_Histo.Text = "H": fgSAB_Histo.CellBackColor = wColor
        fgSAB_Histo.Col = 1: fgSAB_Histo.Text = dateImp10_S(rsSab("SWIHIADEN") + 19000000) & "  " & timeImp8(rsSab("SWIHIAHEN"))
        fgSAB_Histo.CellBackColor = wColor
        
        fgSAB_Histo.Col = 2: fgSAB_Histo.Text = xProcess: fgSAB_Histo.CellBackColor = wColor
        fgSAB_Histo.Col = 3: fgSAB_Histo.Text = rsSab("SWIHIAUTI"): fgSAB_Histo.CellBackColor = wColor
        rsSab.MoveNext
    Loop

End If
'________________________________________________________________________________________________

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub


Public Sub fgSAB_Histo_Display_E()
Dim xSql As String, wColor As Long
Dim blnOk As Boolean, blnDisplay As Boolean
Dim mJOENTT As String, mJOUSER As String
Dim xProcess As String
On Error GoTo Error_Handler

xSql = "select * from " & paramIBM_Library_SABJRN & ".JSWIENA0 W , " & paramIBM_Library_SABJRN & ".JRNENT0 J" _
     & " Where W.JORCV = J.JORCV And W.JOSEQN = J.JOSEQN" _
     & " and SWIENAINT = " & oldYSWISAB0.SWISABZSWI _
     & " order by W.JORCV , W.JOSEQN"
blnOk = False
Set rsSab = cnsab.Execute(xSql)
If Not rsSab.EOF Then
'________________________________________________________________________________________________
    Do While Not rsSab.EOF
        mJOENTT = rsSab("JOENTT")
        wColor = RGB(255, 255, 255)
        blnDisplay = False
        
        Select Case mJOENTT
            Case "PX": xProcess = "importé": blnDisplay = True
            Case "UP": xProcess = "modifié"
                       If mJOUSER <> rsSab("JOUSER") Then blnDisplay = True: wColor = mColor_Y1
            Case "DL":
                If rsSab("SWIENACET") = " " Then
                    xProcess = "traité": wColor = mColor_G1: blnDisplay = True
                Else
                    xProcess = "supprimé": wColor = mColor_W0: blnDisplay = True
                End If
        End Select
        If blnDisplay Then
            mJOUSER = rsSab("JOUSER")
            fgSAB_Histo.Rows = fgSAB_Histo.Rows + 1
            fgSAB_Histo.Row = fgSAB_Histo.Rows - 1
            fgSAB_Histo.Col = 0: fgSAB_Histo.Text = "J": fgSAB_Histo.CellBackColor = wColor
            fgSAB_Histo.Col = 1: fgSAB_Histo.Text = Format(rsSab("JODATE"), "00/00/00") & "  " & Format(rsSab("JOTIME"), "00:00:00")
            fgSAB_Histo.CellBackColor = wColor
            fgSAB_Histo.Col = 2: fgSAB_Histo.Text = xProcess: fgSAB_Histo.CellBackColor = wColor
            fgSAB_Histo.Col = 3: fgSAB_Histo.Text = mJOUSER: fgSAB_Histo.CellBackColor = wColor
        End If
        rsSab.MoveNext
    Loop
'________________________________________________________________________________________________
Else
    xSql = "select * from " & paramIBM_Library_SAB & ".ZSWIMEA0" _
         & " Where SWIMEANUM = " & oldYSWISAB0.SWISABZSWI
    Set rsSab = cnsab.Execute(xSql)
    
    Do While Not rsSab.EOF
        
        xProcess = "traité": wColor = mColor_Y1
            
        fgSAB_Histo.Rows = fgSAB_Histo.Rows + 1
        fgSAB_Histo.Row = fgSAB_Histo.Rows - 1
        fgSAB_Histo.Col = 0: fgSAB_Histo.Text = "H": fgSAB_Histo.CellBackColor = wColor
        fgSAB_Histo.Col = 1: fgSAB_Histo.Text = dateImp10_S(rsSab("SWIMEADTR") + 19000000)
        fgSAB_Histo.CellBackColor = wColor
        
        fgSAB_Histo.Col = 2: fgSAB_Histo.Text = xProcess: fgSAB_Histo.CellBackColor = wColor
        fgSAB_Histo.Col = 3: fgSAB_Histo.Text = rsSab("SWIMEAUTI"): fgSAB_Histo.CellBackColor = wColor
        rsSab.MoveNext
    Loop

End If
'________________________________________________________________________________________________

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub




