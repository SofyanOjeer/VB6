Attribute VB_Name = "rsZDWHOPE0"
'---------------------------------------------------------
Option Explicit
Type typeZDWHOPE0

  DWHOPEDTX   As Long        'DATE EXTRACTION                       1       8      0     S  *
  DWHOPEETA   As Long        'ETABLISSEMENT                         9       4      0     S  *
  DWHOPEAGE   As Long        'AGENCE                               13       4      0     S  *
  DWHOPESER   As String * 2  'SERVICE                              17       2            A  *
  DWHOPESSE   As String * 2  'SOUS-SERVICE                         19       2            A  *
  DWHOPEOPR   As String * 6  'CODE OPERATION                       21       6            A  *
  DWHOPENAT   As String * 6  'CODE NATURE                          27       6            A  *
  DWHOPENDO   As Long        'NUMERO DOSSIER                       33       9      0     S  *
  DWHOPESEQ   As Long        'SEQUENCE                             42       7      0     S  *
  DWHOPECON   As String * 7  'CONTREPARTIE                         49       7            A  *
  DWHOPEPAS   As String * 1  'CLIEN PASSAGE/TIERS                  56       1            A  *
  DWHOPEBDF   As String * 3  'CODE B.D.F                           57       3            A  *
  DWHOPECRE   As Long        'DATE CREATION                        60       8      0     S  *
  DWHOPEENG   As Long        'DATE ENGAGEMENT                      68       8      0     S  *
  DWHOPEDIS   As Long        'DATE MISE A DISPOS                   76       8      0     S  *
  DWHOPEFIN   As Long        'DATE DE FIN                          84       8      0     S  *
  DWHOPEDEV   As String * 3  'DEVISE                               92       3            A  *
  DWHOPEMON   As Currency    'MONTANT                              95      18      3     S  *
  DWHOPEVAL   As Currency    'CONTREVALEUR                        113      18      3     S  *
  DWHOPECOE   As String * 3  'CODE ETAT                           131       3            A  *
  DWHOPECDA   As String * 3  'CODE AUTORISATION                   134       3            A  *
  DWHOPENOA   As String * 6  'N°  AUTORISATION                    137       6            A  *
  DWHOPEDEA   As String * 3  'DEVISE AUTORISATION                 143       3            A  *
  DWHOPEAUT   As String * 20 'COMPTE AUTORISATION                 146      20            A  *
  DWHOPEDUR6  As Long        '  EN JOURS                          166       6      0     S  *
  DWHOPERES6  As Long        '  EN JOURS                          172       6      0     S  *
  DWHOPETYP   As String * 2  'TYPE D'OPÉRATION                    178       2            A  *
  DWHOPEFIX   As Double      'FINXING DATE ARRÊTÉ                 180      14      9     S  *
  DWHOPENOUV  As String * 1  'NOUVELLE OPERATION                  194       1            A  *
  DWHOPEMOIN  As Currency    'MONTANT INITIAL                     195      18      3     S  *
  DWHOPECMOI  As Currency    'CV MONTANT INITIAL                  213      18      3     S  *
  DWHOPEMIN   As Currency    'MT INTÉRÊTS COURUS                  231      18      3     S  *
  DWHOPEVIN   As Currency    'CTVL INT.COURUS                     249      18      3     S  *
  DWHOPEMFI   As Currency    'COURUS DU MOIS                      267      18      3     S  *
  DWHOPECFI   As Currency    'CTVL COURUS MOIS                    285      18      3     S  *
  DWHOPEDSY   As Long        'DATE SYSTÈME                        303       8      0     S  *

End Type

'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsZDWHOPE0_GetBuffer(rsSab As ADODB.Recordset, rsZDWHOPE0 As typeZDWHOPE0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsZDWHOPE0_GetBuffer = Null


rsZDWHOPE0.DWHOPEDTX = rsSab("DWHOPEDTX")
rsZDWHOPE0.DWHOPEETA = rsSab("DWHOPEETA")
rsZDWHOPE0.DWHOPEAGE = rsSab("DWHOPEAGE")
rsZDWHOPE0.DWHOPESER = rsSab("DWHOPESER")
rsZDWHOPE0.DWHOPESSE = rsSab("DWHOPESSE")
rsZDWHOPE0.DWHOPEOPR = rsSab("DWHOPEOPR")
rsZDWHOPE0.DWHOPENAT = rsSab("DWHOPENAT")
rsZDWHOPE0.DWHOPENDO = rsSab("DWHOPENDO")
rsZDWHOPE0.DWHOPESEQ = rsSab("DWHOPESEQ")
rsZDWHOPE0.DWHOPECON = rsSab("DWHOPECON")
rsZDWHOPE0.DWHOPEPAS = rsSab("DWHOPEPAS")
rsZDWHOPE0.DWHOPEBDF = rsSab("DWHOPEBDF")
rsZDWHOPE0.DWHOPECRE = rsSab("DWHOPECRE")
rsZDWHOPE0.DWHOPEENG = rsSab("DWHOPEENG")
rsZDWHOPE0.DWHOPEDIS = rsSab("DWHOPEDIS")
rsZDWHOPE0.DWHOPEFIN = rsSab("DWHOPEFIN")
rsZDWHOPE0.DWHOPEDEV = rsSab("DWHOPEDEV")
rsZDWHOPE0.DWHOPEMON = rsSab("DWHOPEMON")
rsZDWHOPE0.DWHOPEVAL = rsSab("DWHOPEVAL")
rsZDWHOPE0.DWHOPECOE = rsSab("DWHOPECOE")
rsZDWHOPE0.DWHOPECDA = rsSab("DWHOPECDA")
rsZDWHOPE0.DWHOPENOA = rsSab("DWHOPENOA")
rsZDWHOPE0.DWHOPEDEA = rsSab("DWHOPEDEA")
rsZDWHOPE0.DWHOPEAUT = rsSab("DWHOPEAUT")
rsZDWHOPE0.DWHOPEDUR6 = rsSab("DWHOPEDUR6")
rsZDWHOPE0.DWHOPERES6 = rsSab("DWHOPERES6")
rsZDWHOPE0.DWHOPETYP = rsSab("DWHOPETYP")
rsZDWHOPE0.DWHOPEFIX = rsSab("DWHOPEFIX")
rsZDWHOPE0.DWHOPENOUV = rsSab("DWHOPENOUV")
rsZDWHOPE0.DWHOPEMOIN = rsSab("DWHOPEMOIN")
rsZDWHOPE0.DWHOPECMOI = rsSab("DWHOPECMOI")
rsZDWHOPE0.DWHOPEMIN = rsSab("DWHOPEMIN")
rsZDWHOPE0.DWHOPEVIN = rsSab("DWHOPEVIN")
rsZDWHOPE0.DWHOPEMFI = rsSab("DWHOPEMFI")
rsZDWHOPE0.DWHOPECFI = rsSab("DWHOPECFI")
rsZDWHOPE0.DWHOPEDSY = rsSab("DWHOPEDSY")



Exit Function

Error_Handler:

rsZDWHOPE0_GetBuffer = Error

End Function


