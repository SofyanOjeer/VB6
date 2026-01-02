Attribute VB_Name = "rsElpKmLink"
'---------------------------------------------------------
Option Explicit
Type typeElpKmLink
    Method                  As String * 12

    ElpKMSrc_Id             As Long
    ElpKMInfo_Id            As String * 20
    Id                      As String * 20
    Pass                    As Long
    Document_Extension      As String * 3
    Document_Id             As Variant
    Memo                    As Variant
    
    
End Type

'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsElpKmLink_GetBuffer(rsAdo As ADODB.Recordset, rsElpKmLink As typeElpKmLink)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsElpKmLink_GetBuffer = Null

rsElpKmLink.Method = ""

rsElpKmLink.ElpKMSrc_Id = rsAdo("ElpKMSrc_Id")
rsElpKmLink.ElpKMInfo_Id = rsAdo("ElpKMInfo_Id")
rsElpKmLink.Id = rsAdo("ID")
rsElpKmLink.Pass = rsAdo("Pass")
rsElpKmLink.Document_Extension = rsAdo("Document_Extension")
rsElpKmLink.Document_Id = rsAdo("Document_Id")
rsElpKmLink.Memo = rsAdo("Memo")

Exit Function

Error_Handler:

rsElpKmLink_GetBuffer = Error

End Function


'







