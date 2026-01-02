Attribute VB_Name = "impressions_xlsManual"
Option Explicit

Public csvManual As Boolean
Public csvFic As Long

Public xlsManual As Boolean
Public xlsForeColor As Long
Public xlsBackColor As Long
Public xlsRed As Long
Public xlsBlue As Long
Public xlsBlack As Long
Public xlsWhite As Long
Public xlsCol As Long
Public xlsFontSize As Long
Public xlsBackEntete As Long
Public xlsGray As Long

Type TypePageSetup
        PrintArea As String
        LeftFooter As String
        RightFooter As String
        LeftMargin As Double
        RightMargin As Double
        TopMargin As Double
        BottomMargin As Double
        HeaderMargin As Double
        FooterMargin As Double
        Orientation As Long
        Zoom As Long
End Type
Public zoneImpressionPagesetup As TypePageSetup
Public appExcelPublic As Excel.Application
Public Sub ECRIT_LOG2008(mes As String)
Dim fic As Long

    'la log est désactivée, le 16/08/2019
    Exit Sub
    mes = nomDuServeur & " - " & mes
    fic = FreeFile
    Open "C:\Temp\AUTO_JRN.log" For Append As #fic
    Print #fic, mes & Format(Now, "dd/MM/yyyy  HH:nn:ss")
    Close #fic
    
End Sub

Public Function convertitEnMillisecondes(fz As String) As Long
Dim s() As String
Dim SH As Long
Dim sm As Long
Dim ss As Long
Dim sms As Long
    
    s = Split(fz, ":")
    SH = s(0) * 3600
    sm = s(1) * 60
    ss = s(2)
    sms = s(3)
    convertitEnMillisecondes = ((SH + sm + ss) * 1000) + sms

End Function









Public Function indice_feuille_modele(modele As String, wbExcel As Excel.Workbook) As Long
Dim ii As Long
Dim retour As Long

    retour = 0
    For ii = 1 To wbExcel.Sheets.Count
        If wbExcel.Sheets(ii).Name = modele Then
            retour = ii
            Exit For
        End If
    Next ii
    indice_feuille_modele = retour

End Function

Public Function indice_nouvelle_feuille(wbExcel As Excel.Workbook) As Long
Dim ii As Long
Dim retour As Long

    retour = 0
    For ii = 1 To wbExcel.Sheets.Count
        If Left(wbExcel.Sheets(ii).Name, 5) = "Feuil" Or Left(wbExcel.Sheets(ii).Name, 5) = "Sheet" Then
            retour = ii
            Exit For
        End If
    Next ii
    indice_nouvelle_feuille = retour
    
End Function

Public Sub init_TypePagesetup()

        zoneImpressionPagesetup.PrintArea = ""
        zoneImpressionPagesetup.LeftFooter = ""
        zoneImpressionPagesetup.RightFooter = ""
        'zoneImpressionPagesetup.LeftMargin = Application.InchesToPoints(0.25)
        'zoneImpressionPagesetup.RightMargin = Application.InchesToPoints(0.25)
        'zoneImpressionPagesetup.TopMargin = Application.InchesToPoints(0.75)
        'zoneImpressionPagesetup.BottomMargin = Application.InchesToPoints(0.75)
        zoneImpressionPagesetup.HeaderMargin = Application.InchesToPoints(0.3)
        zoneImpressionPagesetup.FooterMargin = Application.InchesToPoints(0.3)
        zoneImpressionPagesetup.Orientation = xlPortrait
        zoneImpressionPagesetup.Zoom = 100

End Sub


Public Sub init_xlsManual()

    xlsBackEntete = RGB(243, 250, 255)
    xlsBlue = RGB(13, 87, 155)
    xlsBlack = RGB(0, 0, 0)
    xlsWhite = RGB(255, 255, 255)
    xlsRed = RGB(192, 0, 0)
    xlsGray = RGB(242, 242, 242)

End Sub
Public Function insere_entete_page(ByRef wsExcel As Excel.Worksheet, ligneDebutFin As String, nbLig As Long, ou As Long) As Long

    wsExcel.Activate
    wsExcel.HPageBreaks.Add Before:=Rows(CStr(ou + 1) & ":" & CStr(ou + 1))
    Rows(ligneDebutFin).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range(CStr(ou + 1) & ":" & CStr(ou + 1)).Select
    ActiveSheet.Paste
    insere_entete_page = nbLig
    
End Function




Public Function nomDuServeur() As String
Dim X As String
Dim I As Long

    nomDuServeur = ""
    X = Space(255)
    Call GetComputerName(X, Len(X))
    For I = 1 To Len(Trim(X))
        If Mid(X, I, 1) <> "" And Asc(Mid(X, I, 1)) <> 0 Then
            nomDuServeur = nomDuServeur & Mid(X, I, 1)
        End If
    Next I
    nomDuServeur = Trim(UCase(nomDuServeur))
    'nomDuServeur = "BIA2008"
    
End Function

Public Sub prtIMP_PDF_Monitor_xlsManual()
Dim currentFileName As String
Dim tmpFileName As String
Static prtIMP_PDF_Seq As Long

        currentFileName = paramEditionNoPaper_Auto_PgmName
        tmpFileName = paramIMP_PDF_Path & "\" & currentFileName & ".pdf"
        prtIMP_PDF_Seq = prtIMP_PDF_Seq + 1
        If Mid$(prtPgmName, 1, 2) = "\\" Then
            prtIMP_PDF_FileName = prtPgmName
        Else
            If blnEditionNoPaper_Auto Then
                Dim xUnit As String, xDir_Save As String, xFile_Save As String
                xDir_Save = paramEditionNoPaper_Folder & "PDF\" & paramEditionNoPaper_Auto_Dir & "_" & YBIATAB0_DATE_CPT_J
                If Not msFileSystem.FolderExists(xDir_Save) Then MkDir xDir_Save

                xUnit = Table_Unit_SSI("S", paramEditionNoPaper_Auto_Unit)
                If xUnit = "" Then xUnit = "S00"
                xFile_Save = xUnit & "." & DSys & "_" & time_Hms & "_" & paramEditionNoPaper_Auto_PgmName & "_" & prtIMP_PDF_Seq & " (" & paramEditionNoPaper_Auto_Unit & ").pdf"
                
                prtIMP_PDF_FileName = xDir_Save & "\_" & xFile_Save
                
                paramEditionNoPaper_Auto_Lnk = "<span style='font-size:9.0pt;font-family:Calibri'>""" _
                                         & "<A HREF=" & Asc34 & Replace(prtIMP_PDF_FileName, paramEditionNoPaper_Folder & "PDF\", paramEditionNoPaper_Partage) & Asc34 & ">" _
                                        & "Cliquez ici pour afficher le document : " & xFile_Save & "</A><BR><BR>"
                                        

            Else
                prtIMP_PDF_FileName = paramIMP_PDF_Path & "\Archive\" & DSys & "_" & time_Hms & "_" & prtIMP_PDF_Seq & "_" & prtPgmName & ".pdf"
            End If
            msFileSystem.MoveFile tmpFileName, prtIMP_PDF_FileName
        End If

End Sub

Public Sub prtSAB_Compta_Mt_xlsManual(lMonTant As Currency, lcolDb As Integer, lcolCr As Integer, ByRef wsExcel As Excel.Worksheet, ByRef currentRow As Long, fSize As Long)
Dim xStr As String
Dim laColonne As Long

    xStr = Format$(Abs(lMonTant), "## ### ### ### ### ##0.00")
    laColonne = IIf(lMonTant < 0, lcolCr, lcolDb)
    'wsexcel.Cells(currentrow, laColonne).Font.Bold = True
    'wsexcel.Cells(currentrow, laColonne).Font.Size = fSize
    'wsexcel.Cells(currentrow, laColonne).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
    wsExcel.Cells(currentRow, laColonne) = xStr

End Sub

Public Function retourne_fin_de_sheet(wsExcel As Excel.Worksheet) As Long
Dim Trouve As Range
Dim PlageDeRecherche As Range
Dim retour As Long

    wsExcel.Activate
    Set PlageDeRecherche = wsExcel.Columns(1)
    Set Trouve = PlageDeRecherche.Find(what:="END_OF_SHEET", LookAt:=xlWhole)
    If Trouve Is Nothing Then
        retour = -1
    Else
        retour = Trouve.Row - 1
    End If
    retourne_fin_de_sheet = retour
    
End Function








Public Sub SetTypePageSetup(wsExcel As Excel.Worksheet)

    With wsExcel.PageSetup
        .PrintArea = CStr(zoneImpressionPagesetup.PrintArea)
        .LeftFooter = zoneImpressionPagesetup.LeftFooter
        .RightFooter = zoneImpressionPagesetup.RightFooter
        .LeftMargin = zoneImpressionPagesetup.LeftMargin
        .RightMargin = zoneImpressionPagesetup.RightMargin
        .TopMargin = zoneImpressionPagesetup.TopMargin
        .BottomMargin = zoneImpressionPagesetup.BottomMargin
        .HeaderMargin = zoneImpressionPagesetup.HeaderMargin
        .FooterMargin = zoneImpressionPagesetup.FooterMargin
        .Orientation = zoneImpressionPagesetup.Orientation
        .Zoom = zoneImpressionPagesetup.Zoom
    End With
End Sub

