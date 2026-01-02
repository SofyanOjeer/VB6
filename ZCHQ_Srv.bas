Attribute VB_Name = "srvZCHQ"
Option Explicit

Public Function ZCHQHIS0_Sql(lCHQHISCPT As String) As String
Dim wCHQCOMREM As Long, xSql As String
Dim K As Integer
Dim Nb As Integer, Nb0 As Integer
Set rsSab = Nothing

xSql = "select CHQhisREM , CHQHISDSP from " & paramIBM_Library_SAB & ".ZCHQHIS0 where CHQHISCOM = '" & lCHQHISCPT & "'"
Set rsSab = cnsab.Execute(xSql)
Nb = 0: Nb0 = 0
If rsSab.EOF Then
    ZCHQHIS0_Sql = "NEANT"
Else
    Do While Not rsSab.EOF
        If Val(rsSab("CHQHISDSP")) = 0 Then
            wCHQCOMREM = rsSab("CHQHISREM")
            If wCHQCOMREM = 0 Then
                Nb0 = Nb0 + 1
            Else
                Nb = Nb + 1
            End If
        End If
        rsSab.MoveNext
    Loop
    ZCHQHIS0_Sql = Nb & " remis"
    If Nb0 > 0 Then ZCHQHIS0_Sql = ZCHQHIS0_Sql & " / " & Nb0 & " non remis"
End If


End Function

Public Function ZCHQDEM0_Sql(lCHQDEMCPT As String) As String
Dim Nb As Integer, xSql As String
Set rsSab = Nothing

xSql = "select count(*) as Tally from " & paramIBM_Library_SAB & ".ZCHQDEM0 where CHQDEMCOM = '" & lCHQDEMCPT & "'"
Set rsSab = cnsab.Execute(xSql)
Nb = rsSab("Tally")
If Nb = 0 Then
    ZCHQDEM0_Sql = "NEANT"
Else
    ZCHQDEM0_Sql = Nb & "carnet(s)"
End If


End Function


Public Function ZCHQCOM0_Sql(lCHQCOMCPT As String) As String
Dim wCHQCOMDT1 As Long, xSql As String
Dim Nb As Integer, Nb0 As Integer
Set rsSab = Nothing

xSql = "select CHQCOMDT1 from " & paramIBM_Library_SAB & ".ZCHQCOM0 where CHQCOMCOM = '" & lCHQCOMCPT & "'"
Set rsSab = cnsab.Execute(xSql)

Nb = 0: Nb0 = 0
If rsSab.EOF Then
    ZCHQCOM0_Sql = "NEANT"
Else
    Do While Not rsSab.EOF
        wCHQCOMDT1 = rsSab("CHQCOMDT1")
        If wCHQCOMDT1 = 0 Then
            Nb0 = Nb0 + 1
        Else
            Nb = Nb + 1
        End If
        rsSab.MoveNext
    
    Loop
    ZCHQCOM0_Sql = Nb & " en cours"
    If Nb0 > 0 Then ZCHQCOM0_Sql = ZCHQCOM0_Sql & " / " & Nb0 & " supprimé(s)"

End If

End Function

