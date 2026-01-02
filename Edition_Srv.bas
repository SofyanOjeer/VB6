Attribute VB_Name = "srvEdition"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const constEdition_Form = "Edition_Form"

Type typeSplfJob
 '   Obj                     As String * 12
 '   Method                  As String * 12
 '   Err                     As String * 10
    
    SJQAMJ                  As String * 8
    SJQID                   As Long
    SJQSEQ                  As Long
    SJQFILE                 As String * 10
    SJQUSR                  As String * 10
    SJQREF                  As String * 10
    SJQSTA                  As String * 3
    SJQPAGENB               As Long
    SJQEXNB                 As Long
    SJQHMS                  As String * 6
    SJQNAME                 As String * 10
    SJQOUTQ                 As String * 10
    SJQXAMJ                 As String * 8
    SJQXHMS                 As String * 6
    SJQXOUTQ                As String * 10
    SJQXSTA                 As String * 3
    SJQXEVTID               As Long
End Type
    
Type typeEdition_Form
 '   Obj                     As String * 12
 '   Method                  As String * 12
 '   Err                     As String * 10
    
    K1              As String * 12
    K2              As String * 12
    Name            As String * 40
    Courrier        As String * 1
    Orientation     As String * 1
    LinePerPage      As Integer
    FontSize        As Integer
    Duplex          As String * 1
    Filigrane       As String * 1
    Copies          As Integer
    PaperBin        As String * 1
    Hold            As String * 1
    Save            As String * 1
    FontName        As String * 30
    PrinterUnit     As String * 1           ' impression poste utilisateur sinon imp réseau  du service
    Unit            As String * 10
'$JPL 2014-12-15
    NoPaper_Prod    As String * 1
    'Unit2           As String * 10          ' service destintaire d'une copie
    'Unit3           As String * 10          ' service destintaire d'une copie
End Type
Type typeEdition
    Obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10

    Nature                  As String * 3
    Id                      As String * 20
    Position                As Long
    Memo1                   As Variant
    Memo2                   As Variant

End Type

Type typeEdition_Memo2
    Length                 As String * 6
    FormatCode             As String * 1
    Format                 As String * 10
    textCode               As String * 1
    Text                   As Variant

End Type

Public frmRTF_Form_K2 As String
Public frmRTF_FileName As String
Public frmRTF_recEdition As typeEdition
Public frmRTF_Caller As String
Public frmRTF_Buffer_Name As String

Public frmRTF_blnOK As Boolean
Public frmRTF_blnCourrier As Boolean
Public frmRTF_prtOrientation As Integer
Public frmRTF_prtPaperSize As Integer

Public mHtml_Head As String
Public Sub recEdition_Init(recEdition As typeEdition)
recEdition.Method = ""
recEdition.Obj = "Edition"
recEdition.Err = ""
recEdition.Id = ""
recEdition.Nature = ""
recEdition.Position = 0
recEdition.Memo1 = ""
recEdition.Memo2 = ""
End Sub

Public Function paramEdition_Init() '(lstErr As ListBox, lcmdContext As CommandButton)
Dim K As Integer, K1 As Integer, X As String
Dim xName As String, xMemo As String
Dim V
On Error GoTo Error_Handler

paramEdition_Init = Null

App_Debug = "paramEdition_Init"
'----------------------------------
V = rsElpTable_Read("Edition", "NoPaper", "Folder", xName, xMemo)
If Not IsNull(V) Then GoTo Error_MsgBox
paramEditionNoPaper_Folder = paramServer(xMemo)

paramEditionNoPaper_Folder_MakePDF = paramEditionNoPaper_Folder & "MakePDF\"

V = rsElpTable_Read("Edition", "NoPaper", "Partage", xName, xMemo)
If Not IsNull(V) Then GoTo Error_MsgBox
paramEditionNoPaper_Partage = "\\" & Trim(xMemo)

V = rsElpTable_Read("Edition", "Splf", "Folder", xName, xMemo)
If Not IsNull(V) Then GoTo Error_MsgBox
paramEditionSplf_Folder = paramServer(xMemo)

V = rsElpTable_Read("Edition", "Filigrane", "Folder", xName, xMemo)
If Not IsNull(V) Then GoTo Error_MsgBox
paramEditionFiligrane_Folder = paramServer(xMemo)

V = rsElpTable_Read("Edition", "Courrier", "Folder", xName, xMemo)
If Not IsNull(V) Then GoTo Error_MsgBox
paramEditionCourrier_Folder = paramServer(xMemo)

V = rsElpTable_Read("Edition", "Ftp", "File", xName, paramEditionFtp_File)
If Not IsNull(V) Then GoTo Error_MsgBox

V = rsElpTable_Read("Edition", "Archive", "Folder", xName, xMemo)
If Not IsNull(V) Then GoTo Error_MsgBox
paramEditionArchive_Folder = paramServer(xMemo)

V = rsElpTable_Read("Edition", "Corbeille", "Folder", xName, xMemo)
If Not IsNull(V) Then GoTo Error_MsgBox
paramEditionCorbeille_Folder = paramServer(xMemo)

If blnOff_Line Then
    paramEditionFtp_File = "C:\Temp\S820i_Out\SPLF\SPLFFTPW0"
    paramEditionSplf_Folder = "C:\Temp\Splf\"
    paramEditionCourrier_Folder = "C:\Temp\Splf\Courrier\"
    paramEditionFiligrane_Folder = "C:\Temp\Filigrane\"
    paramEditionArchive_Folder = "C:\Temp\Splf\Archive\"
    paramEditionCorbeille_Folder = "C:\Temp\Splf\Corbeille\"
End If

'--------------------------------------------------------------------------------
If mHtml_Head = "" Then
    Dim msFile As Scripting.File, msFile_rtf As Scripting.TextStream
    X = paramEditionFiligrane_Folder & "VB_HTML_Head.txt"
    Set msFile = msFileSystem.GetFile(X)
    Set msFile_rtf = msFile.OpenAsTextStream(ForReading)
    mHtml_Head = msFile_rtf.ReadAll
End If

blnMakePDF_Actif = True

Exit Function

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
End Function






Public Sub rsEdition_Form_Init(lEdition_Form As typeEdition_Form)
'lEdition_Form.Obj = "ZMNUETA0_S"
'lEdition_Form.Method = ""
'lEdition_Form.Err = ""
lEdition_Form.K1 = ""
lEdition_Form.K2 = ""
lEdition_Form.Name = ""
lEdition_Form.Courrier = "0"
lEdition_Form.FontSize = 6
lEdition_Form.Orientation = "1"
lEdition_Form.Duplex = "1"
lEdition_Form.Filigrane = "A"
lEdition_Form.Copies = 1
lEdition_Form.PaperBin = 7 '2
lEdition_Form.Hold = "0"
lEdition_Form.Save = "0"
lEdition_Form.FontName = prtFontName_CourierNew
lEdition_Form.PrinterUnit = "0"
lEdition_Form.Unit = ""
lEdition_Form.NoPaper_Prod = "0"
'lEdition_Form.Unit2 = ""
'lEdition_Form.Unit3 = ""


End Sub

'---------------------------------------------------------
Public Sub rsEdition_Form_PutBuffer(lMemo As String, recEdition_Gestion As typeEdition_Form)
'---------------------------------------------------------

Mid$(lMemo, 1, 1) = recEdition_Gestion.Courrier
Mid$(lMemo, 2, 1) = recEdition_Gestion.Orientation
Mid$(lMemo, 3, 3) = Format$(recEdition_Gestion.LinePerPage, "000")
Mid$(lMemo, 6, 2) = Format$(recEdition_Gestion.FontSize, "00")
Mid$(lMemo, 8, 1) = recEdition_Gestion.Duplex
Mid$(lMemo, 9, 1) = recEdition_Gestion.Filigrane
Mid$(lMemo, 10, 2) = Format$(recEdition_Gestion.Copies, "00")
Mid$(lMemo, 12, 1) = recEdition_Gestion.PaperBin
Mid$(lMemo, 13, 1) = recEdition_Gestion.Hold
Mid$(lMemo, 14, 1) = recEdition_Gestion.Save
Mid$(lMemo, 15, 30) = recEdition_Gestion.FontName
Mid$(lMemo, 45, 1) = recEdition_Gestion.PrinterUnit
Mid$(lMemo, 46, 10) = recEdition_Gestion.Unit
'$JPL 2014-12-15
Mid$(lMemo, 56, 1) = recEdition_Gestion.NoPaper_Prod
'Mid$(lMemo, 56, 10) = recEdition_Gestion.Unit2
'Mid$(lMemo, 66, 10) = recEdition_Gestion.Unit3

End Sub

'---------------------------------------------------------
Public Function rsEdition_Form_GetBuffer(lMemo As String, recEdition_Gestion As typeEdition_Form)
'---------------------------------------------------------
rsEdition_Form_GetBuffer = Null

    recEdition_Gestion.Courrier = Mid$(lMemo, 1, 1)
    recEdition_Gestion.Orientation = Mid$(lMemo, 2, 1)
    recEdition_Gestion.LinePerPage = CInt(Mid$(lMemo, 3, 3))
    recEdition_Gestion.FontSize = CInt(Mid$(lMemo, 6, 2))
    recEdition_Gestion.Duplex = Mid$(lMemo, 8, 1)
    recEdition_Gestion.Filigrane = Mid$(lMemo, 9, 1)
    recEdition_Gestion.Copies = CInt(Mid$(lMemo, 10, 2))
    recEdition_Gestion.PaperBin = Mid$(lMemo, 12, 1)
    recEdition_Gestion.Hold = Mid$(lMemo, 13, 1)
    recEdition_Gestion.Save = Mid$(lMemo, 14, 1)
    recEdition_Gestion.FontName = Mid$(lMemo, 15, 30)
    recEdition_Gestion.PrinterUnit = Mid$(lMemo, 45, 1)
    recEdition_Gestion.Unit = Mid$(lMemo, 46, 10)
'$JPL 2014-12-15
    recEdition_Gestion.NoPaper_Prod = Val(Mid$(lMemo, 56, 1))
    
    'recEdition_Gestion.Unit2 = Mid$(lMemo, 56, 10)
    'recEdition_Gestion.Unit3 = Mid$(lMemo, 66, 10)
    

End Function

Public Function rsEdition_Form(lEdition_Form As typeEdition_Form) As String
Dim V, xMemo As String
V = rsElpTable_Read(constEdition_Form, lEdition_Form.K1, lEdition_Form.K2, lEdition_Form.Name, xMemo)
If IsNull(V) Then
    rsEdition_Form_GetBuffer xMemo, lEdition_Form
Else
    rsEdition_Form_Init lEdition_Form
End If
rsEdition_Form = lEdition_Form.Name
End Function

