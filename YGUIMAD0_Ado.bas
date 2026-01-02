Attribute VB_Name = "adoYGUIMAD0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsYGUIMAD0_PutBuffer(rsADO As ADODB.Recordset, rsYGUIMAD0 As typeYGUIMAD0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsYGUIMAD0_PutBuffer = Null
rsADO("GUIMADID") = rsYGUIMAD0.GUIMADID
rsADO("GUIESPOPE") = rsYGUIMAD0.GUIESPOPE
rsADO("GUIESPDOS") = rsYGUIMAD0.GUIESPDOS
rsADO("GUIESPNAT") = rsYGUIMAD0.GUIESPNAT
rsADO("GUIESPMON") = rsYGUIMAD0.GUIESPMON
rsADO("GUIESPDEV") = rsYGUIMAD0.GUIESPDEV
rsADO("GUIESPCP1") = rsYGUIMAD0.GUIESPCP1
rsADO("GUIESPCL1") = rsYGUIMAD0.GUIESPCL1
rsADO("GUIESPTI1") = rsYGUIMAD0.GUIESPTI1
rsADO("GUIESPDJO") = rsYGUIMAD0.GUIESPDJO

rsADO("GUIMADMON") = rsYGUIMAD0.GUIMADMON
rsADO("GUIMADTDO") = rsYGUIMAD0.GUIMADTDO
rsADO("GUIMADTIN") = rsYGUIMAD0.GUIMADTIN
rsADO("GUIMADMOT") = rsYGUIMAD0.GUIMADMOT
rsADO("GUIMADLIEN") = rsYGUIMAD0.GUIMADLIEN
rsADO("GUIMADSTA") = rsYGUIMAD0.GUIMADSTA

rsADO("GUIMADUPDS") = rsYGUIMAD0.GUIMADUPDS
    
Exit Function

Error_Handler:

rsYGUIMAD0_PutBuffer = Error

End Function


'---------------------------------------------------------
Public Function adoYGUIMAD0_AddNew(rsADO As ADODB.Recordset, rsYGUIMAD0 As typeYGUIMAD0)
'---------------------------------------------------------
On Error GoTo Error_Handler

adoYGUIMAD0_AddNew = Null
rsADO.AddNew
adoYGUIMAD0_AddNew = rsYGUIMAD0_PutBuffer(rsADO, rsYGUIMAD0)
rsADO.Update

Exit Function

Error_Handler:

adoYGUIMAD0_AddNew = Error

End Function



