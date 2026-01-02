Attribute VB_Name = "srvDocuShare"
Option Explicit

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Const DSAXES_PROPID_CORE_CUSTOM = &H1
Const DSAXES_PROPID_CORE_TITLE = &H2
Const DSAXES_PROPID_CORE_SUMMARY = &H4
Const DSAXES_PROPID_CORE_DESCRIPTION = &H8
Const DSAXES_PROPID_CORE_KEYWORDS = &H10
Const DSAXES_PROPID_CORE_OWNER = &H40
Const DSAXES_PROPID_CORE_CREATEDATE = &H80
Const DSAXES_PROPID_CORE_MODIFIEDDATE = &H100
Const DSAXES_PROPID_CORE_MODIFIEDBY = &H200
Const DSAXES_PROPID_CORE_PARENTHANDLES = &H400

Const DSCONTF_PARENT = &H1
Const DSCONTF_FOLDERS = &H2
Const DSCONTF_DOCUMENTS = &H4
Const DSCONTF_NONDOCUMENTS = &H8
Const DSCONTF_VERSIONS = &H10
Const DSCONTF_IDENTITY = &H20
Const DSCONTF_DESCENDANTS = &H40
Const DSCONTF_ASCENDANTS = &H80
Const DSCONTF_ASCEND_ROOTFIRST = &H100

Const SVRMAP_OPTION_SILENT = &H4000&
Const SVRMAP_OPTION_NOERRORDLG = &H8000&
'
Const DSAXES_MODE_ENABLEXML = &H80
Const DSAXES_MODE_SAVESETTINGS = &H2000
Const DSAXES_MODE_LOADSETTINGS = &H4000
Const DSAXES_MODE_ADDDSCGI = &H8
'DSAXES_MODE_STDMASK = &H7FFF 'valeur DS2
Const DSAXES_MODE_STDMASK = &HFF 'valeur DS3
Const DSSRCH_DEFAULT_MODE = &H0 'rajouté dans ds3
Const DSSRCH_TYPE_COLL = &H2000000
Const DSSRCH_BY_TITLE = &H10000
Const DSSRCH_BY_CREATTIME = &H1000
Const DSSRCH_SCOPE_COLL = &H8000000
Const DSSRCH_TYPE_FILE = &H3000000
    
Const DSAXES_MODE_MEMORIZEUSER = &H10
Const DSAXES_MODE_ENABLECACHE = &H40
'_________________________________________________________________________
Dim server As New DSServerMap.server
Dim resultats As ItemEnum
Dim objTemp As ItemObj
Dim objTemp2 As ItemObj
Dim Col As ItemEnum
Dim lresult As Long
Dim DSCONTF_CHILDREN
Dim DSCONTF_ALL
'_________________________________________________________________________
Dim paramDocuShare_Server As String
Dim paramDocuShare_Username  As String
Dim paramDocuShare_Password  As String
Public paramDocuShare_Folder As String, paramDocuShare_Folder_Document As String
Public paramDocuShare_Collection_SAB_CDO As Long
Public paramDocuShare_Collection_KYC As Long
Public paramDocuShare_Collection_Informatique As Long
Public paramDocuShare_Collection_SI_Doc As Long

Dim blnDS_Server_Open As Boolean
Public Sub DS_Server_Open()

Dim V, wName As String, wMemo As String
On Error GoTo Exit_sub

If blnDS_Server_Open Then Exit Sub


V = rsElpTable_Read("Docushare", "Server", "", wName, paramDocuShare_Server)
If Not IsNull(V) Then GoTo Error_MsgBox

V = rsElpTable_Read("Docushare", "Username", "", wName, paramDocuShare_Username)
If Not IsNull(V) Then GoTo Error_MsgBox

V = rsElpTable_Read("Docushare", "PasswordX", "", wName, paramDocuShare_Password)
If Not IsNull(V) Then GoTo Error_MsgBox

V = rsElpTable_Read("Docushare", "Folder", "", wName, paramDocuShare_Folder)
If Not IsNull(V) Then GoTo Error_MsgBox

V = rsElpTable_Read("Docushare", "Collection", "SAB_CDO", wName, wMemo)
If Not IsNull(V) Then GoTo Error_MsgBox
paramDocuShare_Collection_SAB_CDO = CLng(wMemo)

V = rsElpTable_Read("Docushare", "Collection", "KYC", wName, wMemo)
If Not IsNull(V) Then GoTo Error_MsgBox
paramDocuShare_Collection_KYC = CLng(wMemo)

V = rsElpTable_Read("Docushare", "Collection", "Informatique", wName, wMemo)
If Not IsNull(V) Then GoTo Error_MsgBox
paramDocuShare_Collection_Informatique = CLng(wMemo)

V = rsElpTable_Read("Docushare", "Collection", "SI_Doc", wName, wMemo)
If Not IsNull(V) Then GoTo Error_MsgBox
paramDocuShare_Collection_SI_Doc = CLng(wMemo)

If Not IsNull(DS_Folder_Monitor) Then GoTo Error_MsgBox

DSCONTF_CHILDREN = DSCONTF_FOLDERS Or DSCONTF_DOCUMENTS Or DSCONTF_NONDOCUMENTS
DSCONTF_ALL = DSCONTF_DESCENDANTS Or DSCONTF_CHILDREN

server.Options = SVRMAP_OPTION_NOERRORDLG Or SVRMAP_OPTION_SILENT
server.DocuShareAddress = paramDocuShare_Server
server.UserName = paramDocuShare_Username
server.Password = paramDocuShare_Password

server.Logon
blnDS_Server_Open = True



Exit Sub

Exit_sub:
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, "DS_Server_Open"

End Sub

Public Function DS_Collection_Load() As Long
   Set resultats = DS_Collection_Qry("bascule", 1305, server)
   ' Set resultats = DS_Collection_Qry("CDR", 137, server)
   ' Debug.Print "Nombre de réponses : " & resultats.Length
    DS_Collection_Load = resultats.Length
    While resultats.NextPos <> 0
         Set objTemp = server.CreateObject(resultats.nextItem.Handle)
        
         'Debug.Print "TROUVE : " & objTemp.Title & " (" & objTemp.Handle & ")"
         'listage du contenu de la collection trouvée
         lresult = objTemp.DSLoadChildren
         Set Col = objTemp.EnumObjects(DSCONTF_CHILDREN)
         Col.Reset
         While Col.NextPos <> 0
            Set objTemp2 = Col.nextItem
           ' Debug.Print "   >" & objTemp2.Title
            'est-ce un document?
            If objTemp2.TypeNum = 3 Then
                ' téléchargement
                objTemp2.DSLoadProps 'on charge les propriétés
                objTemp2.Name = paramDocuShare_Folder_Document & "\" & objTemp2.Handle & "-" & objTemp2.Name
                objTemp2.DSDownload (0)
                ' on ouvre
                ShellExecute XForm.hwnd, "open", objTemp2.Name, "", paramDocuShare_Folder, 1
            End If
         Wend
       
    Wend
    

End Function


Public Function DS_Document_Load(ByVal srctitle As String, ByVal collection As Long) As Long

On Error GoTo Exit_Function
If Not blnDS_Server_Open Then Exit Function
    
    Set resultats = DS_Document_Qry(srctitle, collection, server)
   ' Debug.Print "Nombre de réponses : " & resultats.Length
    DS_Document_Load = resultats.Length
    While resultats.NextPos <> 0
         Set objTemp = server.CreateObject(resultats.nextItem.Handle)
            objTemp.DSLoadProps
         'Debug.Print "TROUVE : " & objTemp.Title & " (" & objTemp.Handle & ")"
         'listage du contenu de la collection trouvée
                objTemp.Name = paramDocuShare_Folder_Document & "\" & objTemp.Handle & "-" & objTemp.Name
                objTemp.DSDownload (0)
                ' on ouvre
                'ShellExecute XForm.hwnd, "print", objTemp.Name, "", paramDocuShare_Folder, 1
                ShellExecute XForm.hwnd, "open", objTemp.Name, "", paramDocuShare_Folder, 1

    Wend
    
Exit_Function:

End Function

Public Function DS_Collection_Qry(ByVal srctitle As String, ByVal collection As Long, objServer) As ItemEnum
'retourne un itemenum
'sinon rien

'nomdossiertmp est la collection recherchée
'collectionderecherche est la collection contenante
'iddossiertrouve est l'id du dossier trouvé
Dim Gateway
Set Gateway = objServer.Open
'Dim result
'Set result = CreateObject("DSITEMENUMLib.EnumObj")

DSCONTF_CHILDREN = (DSCONTF_FOLDERS Or DSCONTF_DOCUMENTS Or DSCONTF_NONDOCUMENTS)
DSCONTF_ALL = (DSCONTF_IDENTITY Or DSCONTF_PARENT Or DSCONTF_CHILDREN)

Dim Status
Status = 0
Dim xmlData
Dim queryDepth
Gateway.DSMaxItems = 100
Gateway.DSCollHandle = collection
Gateway.DSDisplayName = srctitle
    
Status = Gateway.Search(DSSRCH_TYPE_COLL Or DSSRCH_SCOPE_COLL Or DSSRCH_BY_TITLE, "")
Status = Gateway.ServerResponse.ContentData.ParseQueryResults
If Status <> 0 Then
    'Erreur dans la recherche
End If
Dim itemList As ItemEnum
On Error Resume Next
   Set itemList = Gateway.ServerResponse.ContentData.ItemDescriptorArray.Enumerator(DSCONTF_CHILDREN Or DSCONTF_DESCENDANTS)
If Err.Number <> 0 Then
            
End If
        
itemList.Reset
        
Set DS_Collection_Qry = itemList

End Function


Public Function DS_Document_Qry(ByVal srctitle As String, ByVal collection As Long, objServer) As ItemEnum
'retourne un itemenum
'sinon rien

'nomdossiertmp est la collection recherchée
'collectionderecherche est la collection contenante
'iddossiertrouve est l'id du dossier trouvé
Dim Gateway
Set Gateway = objServer.Open
'Dim result
'Set result = CreateObject("DSITEMENUMLib.EnumObj")
Const DSAXES_MODE_MEMORIZEUSER = &H10
Const DSAXES_MODE_ENABLECACHE = &H40

DSCONTF_CHILDREN = (DSCONTF_FOLDERS Or DSCONTF_DOCUMENTS Or DSCONTF_NONDOCUMENTS)
DSCONTF_ALL = (DSCONTF_IDENTITY Or DSCONTF_PARENT Or DSCONTF_CHILDREN)

Dim Status
Status = 0
Dim xmlData
Dim queryDepth
Gateway.DSMaxItems = 100
Gateway.DSCollHandle = collection
Gateway.DSDisplayName = srctitle
    
Status = Gateway.Search(DSSRCH_TYPE_FILE Or DSSRCH_SCOPE_COLL Or DSSRCH_BY_TITLE, "")
Status = Gateway.ServerResponse.ContentData.ParseQueryResults
If Status <> 0 Then
    'Erreur dans la recherche
End If
Dim itemList As ItemEnum
On Error Resume Next
   Set itemList = Gateway.ServerResponse.ContentData.ItemDescriptorArray.Enumerator(DSCONTF_CHILDREN Or DSCONTF_DESCENDANTS)
If Err.Number <> 0 Then
            
End If
        
itemList.Reset
        
Set DS_Document_Qry = itemList

End Function


Public Function DS_Folder_Monitor()
Dim V, K As Integer, X8 As String
Dim fsoFile As File

On Error GoTo Error_Handler
DS_Folder_Monitor = Null
If Not msFileSystem.FolderExists(paramDocuShare_Folder) Then MkDir paramDocuShare_Folder
paramDocuShare_Folder_Document = paramDocuShare_Folder & "\Docushare"
If Not msFileSystem.FolderExists(paramDocuShare_Folder_Document) Then MkDir paramDocuShare_Folder_Document

frmElp.filDoc.Path = paramDocuShare_Folder_Document
frmElp.filDoc.Pattern = "*.*"


For K = 0 To frmElp.filDoc.ListCount - 1
    frmElp.filDoc.ListIndex = K
    Set fsoFile = msFileSystem.GetFile(frmElp.filDoc.Path & "\" & frmElp.filDoc.FileName)
    If Err = 0 Then
        Call dateJMA6_AMJ(fsoFile.DateLastModified, X8)
        If X8 < DSys Then msFileSystem.DeleteFile frmElp.filDoc.Path & "\" & frmElp.filDoc.FileName, True
    End If
    
Next K

Exit Function

Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, "DS_Folder_Monitor"
    DS_Folder_Monitor = V

End Function
