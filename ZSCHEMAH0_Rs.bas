Attribute VB_Name = "rsZSCHEMAH0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Type typeZSCHEMAH0
    SCHEMAETA       As Integer                        ' ETABLISSEMENT
    SCHEMAOPE       As String * 3                     ' TABLES BASE 023
    SCHEMAEVE       As String * 3                     ' TABLES BASE 024
    SCHEMAPLA       As Long                           ' PLAN
    SCHEMAARG       As String * 18                    ' ARGUMENT
    SCHEMANUM       As Long                           ' NUMERO DE LIGNE
    SCHEMASCH       As Integer                        ' CODE SCHEMA
    SCHEMALIB       As String * 30                    ' LIBELLE
    SCHEMADO1       As String                         ' ZONE DE DONNEE
    SCHEMADO2       As String                         ' ZONE DE DONNEE
    SCHEMADO3       As String                         ' ZONE DE DONNEE
    SCHEMADO4       As String                         ' ZONE DE DONNEE
    SCHEMAHDT       As Long                           ' HI-D-DATE MAJ
    SCHEMAHHE       As Long                           ' HI-D-HEURE MAJ
    SCHEMAHUT       As Integer                        ' HI-D-CODE UTILIS
    SCHEMAHDC       As Long                           ' HI-D-DATE COMPTA
    SCHEMAHTY       As String * 1                     ' HI-D-TYPE MAJ
    SCHEMAHOP       As String * 1                     ' HI-D-OPTION MAJ
    SCHEMAFDT       As Long                           ' HI-F-DATE MAJ
    SCHEMAFHE       As Long                           ' HI-F-HEURE MAJ
    SCHEMAFUT       As Integer                        ' HI-F-CODE UTILIS
    SCHEMAFDC       As Long                           ' HI-F-DATE COMPTA
    SCHEMAFTY       As String * 1                     ' HI-F-TYPE MAJ
    SCHEMAFOP       As String * 1                     ' HI-F-OPTION MAJ

End Type
Public Sub rsYSCHEMAH0_Init(rsYSCHEMAH0 As typeZSCHEMAH0)
rsYSCHEMAH0.SCHEMAETA = 0
rsYSCHEMAH0.SCHEMAOPE = ""
rsYSCHEMAH0.SCHEMAEVE = ""
rsYSCHEMAH0.SCHEMAPLA = 0
rsYSCHEMAH0.SCHEMAARG = ""
rsYSCHEMAH0.SCHEMANUM = 0
rsYSCHEMAH0.SCHEMASCH = 0
rsYSCHEMAH0.SCHEMALIB = ""
rsYSCHEMAH0.SCHEMADO1 = ""
rsYSCHEMAH0.SCHEMADO2 = ""
rsYSCHEMAH0.SCHEMADO3 = ""
rsYSCHEMAH0.SCHEMADO4 = ""
rsYSCHEMAH0.SCHEMAHDT = 0
rsYSCHEMAH0.SCHEMAHHE = 0
rsYSCHEMAH0.SCHEMAHUT = 0
rsYSCHEMAH0.SCHEMAHDC = 0
rsYSCHEMAH0.SCHEMAHTY = ""
rsYSCHEMAH0.SCHEMAHOP = ""
rsYSCHEMAH0.SCHEMAFDT = 0
rsYSCHEMAH0.SCHEMAFHE = 0
rsYSCHEMAH0.SCHEMAFUT = 0
rsYSCHEMAH0.SCHEMAFDC = 0
rsYSCHEMAH0.SCHEMAFTY = ""
rsYSCHEMAH0.SCHEMAFOP = ""
End Sub
Public Function rsZSCHEMAH0_GetBuffer(rsADO As ADODB.Recordset, rsZSCHEMAH0 As typeZSCHEMAH0)
On Error GoTo Error_Handler
rsZSCHEMAH0_GetBuffer = Null
rsZSCHEMAH0.SCHEMAETA = rsADO("SCHEMAETA")
rsZSCHEMAH0.SCHEMAOPE = rsADO("SCHEMAOPE")
rsZSCHEMAH0.SCHEMAEVE = rsADO("SCHEMAEVE")
rsZSCHEMAH0.SCHEMAPLA = rsADO("SCHEMAPLA")
rsZSCHEMAH0.SCHEMAARG = rsADO("SCHEMAARG")
rsZSCHEMAH0.SCHEMANUM = rsADO("SCHEMANUM")
rsZSCHEMAH0.SCHEMASCH = rsADO("SCHEMASCH")
rsZSCHEMAH0.SCHEMALIB = rsADO("SCHEMALIB")
rsZSCHEMAH0.SCHEMADO1 = rsADO("SCHEMADO1")
rsZSCHEMAH0.SCHEMADO2 = rsADO("SCHEMADO2")
rsZSCHEMAH0.SCHEMADO3 = rsADO("SCHEMADO3")
rsZSCHEMAH0.SCHEMADO4 = rsADO("SCHEMADO4")
rsZSCHEMAH0.SCHEMAHDT = rsADO("SCHEMAHDT")
rsZSCHEMAH0.SCHEMAHHE = rsADO("SCHEMAHHE")
rsZSCHEMAH0.SCHEMAHUT = rsADO("SCHEMAHUT")
rsZSCHEMAH0.SCHEMAHDC = rsADO("SCHEMAHDC")
rsZSCHEMAH0.SCHEMAHTY = rsADO("SCHEMAHTY")
rsZSCHEMAH0.SCHEMAHOP = rsADO("SCHEMAHOP")
rsZSCHEMAH0.SCHEMAFDT = rsADO("SCHEMAFDT")
rsZSCHEMAH0.SCHEMAFHE = rsADO("SCHEMAFHE")
rsZSCHEMAH0.SCHEMAFUT = rsADO("SCHEMAFUT")
rsZSCHEMAH0.SCHEMAFDC = rsADO("SCHEMAFDC")
rsZSCHEMAH0.SCHEMAFTY = rsADO("SCHEMAFTY")
rsZSCHEMAH0.SCHEMAFOP = rsADO("SCHEMAFOP")
Exit Function
Error_Handler:
rsZSCHEMAH0_GetBuffer = Error
End Function
