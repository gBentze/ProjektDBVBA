# ProjektDBVBA
Option Compare Database


'Tabelle Abteilung, Mitkreis, typeFzgGrp Const dateipfad = "D:\DantenbankAccess\Abschlussphase" Const xlBlattName As String = "Tabelle1"

Private Type typeAbteilung
kuerzel As String
End Type

Private Type typeMitKreis
mitKreis As String
End Type

Private Type typeFzgGp
fzgGrp As String
End Type

Private Type MitarbeiterStruc
dKXKennung As String
nachname As String
vorname As String
email As String
mitKreis As String
lngMitKreis As Long
re As String
stammNr As Long
kstelle As String

End Type

Private AbtXLSXdata() As typeAbteilung
Private MitKreisXLSXdata() As typeMitKreis
Private FzgGrpXLSXdata() As typeFzgGp
Private MitarbeiterXLSXdata() As MitarbeiterStruc
'Anzahl Zeilen in Excel Tabelle

Private XLSXmax As Integer

Private Sub ImportDaten()
Dim xlpfad As String
xlpfad = dateipfad & "\StammdatenListe.xlsx"
Dim vKuerzel As Variant, vMitKreis, vFzgGrp
Dim vDKXKennung As Variant, vNachname, vVorname, vEmail, vMitKreis, vRE, vStammNr, vKostenstelle

Dim i As Integer
Dim iRowS As Integer
Dim iRowL As Integer
Dim iCol As Integer
Dim sCol As String

' Verweis auf Excel-Bibliothek muss gesetzt sein
Dim xlsApp As Excel.Application
Dim Blatt As Excel.Worksheet
Dim MsgAntw As Integer
' Konstante: Name des einzulesenden Arbeitsblattes

' Excel vorbereiten
On Error Resume Next
Set xlsApp = GetObject(, "Excel.Application")
If xlsApp Is Nothing Then
    Set xlsApp = CreateObject("Excel.Application")
End If

On Error GoTo 0
' Exceldatei readonly öffnen
xlsApp.Workbooks.Open xlpfad, , True
    ' Erste Zeile wird statisch angegeben
    iRowS = 2
    ' Letzte Zeile auf Tabellenblatt wird dynamisch ermittelt
    iRowL = xlsApp.Worksheets(xlBlattName).Cells(xlsApp.Rows.Count, 1).End(xlUp).Row
                    
    ' Spalte mit Kuerzel einlesen
    'Stammnummer
    sCol = "A"
    vStammNr = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'Nachname
    sCol = "B"
    vNachname = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'Vorname
    sCol = "C"
    vVorname = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'DKX-Kennung
    sCol = "D"
    vDKXKennung = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'Email
    sCol = "E"
    vEmail = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'Abteilung
    sCol = "F"
    vKuerzel = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'Kostenstelle
    sCol = "G"
    vKostenstelle = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'Rechtseinheit
    sCol = "H"
    vRE = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'Mitarbeiterkreis
    sCol = "J"
    vMitKreis = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
    'FzgGrp
    sCol = "M"
    vFzgGrp = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)

    '--------------------
    ' Werte in globaler Variable speichern
    XLSXmax = iRowL - 1
    ReDim AbtXLSXdata(1 To XLSXmax)
    ReDim MitKreisXLSXdata(1 To XLSXmax)
    ReDim FzgGrpXLSXdata(1 To XLSXmax)
    ReDim MitarbeiterXLSXdata(1 To XLSXmax)
    
    With xlsApp.Worksheets(xlBlattName)
        For i = 1 To XLSXmax
            Debug.Print .Cells(i, 1)
            Debug.Print .Cells(i, 2)
            
            AbtXLSXdata(i).kuerzel = vKuerzel(i, 1)
            MitKreisXLSXdata(i).mitKreis = vMitKreis(i, 1)
            FzgGrpXLSXdata(i).fzgGrp = vFzgGrp(i, 1)
            MitarbeiterXLSXdata(i).dKXKennung = vDKXKennung(i, 1)
            MitarbeiterXLSXdata(i).email = vEmail(i, 1)
            MitarbeiterXLSXdata(i).nachname = vNachname(i, 1)
            MitarbeiterXLSXdata(i).vorname = vVorname(i, 1)
            MitarbeiterXLSXdata(i).mitKreis = vMitKreis(i, 1)
            MitarbeiterXLSXdata(i).re = vRE(i, 1)
            MitarbeiterXLSXdata(i).kstelle = vKostenstelle(i, 1)
            MitarbeiterXLSXdata(i).stammNr = vStammNr(i, 1)

        Next i
    End With

Set xlsApp = Nothing


End Sub

Private Sub WriteXLSXDaten()
' Daten aus Variable in Tabelle übertragen
Dim i As Integer
'Dim lngMitKreisID As Long
Dim sSQLAbt As String
Dim sSQLKreis As String
Dim sSQLGrp As String
Dim sSQLMit As String
Dim sSQLStNr As String
Dim sSQLKst As String

Dim abtID As Long, grpID As Long, kreisID As Long, mitID As Long, reID As Long, StammNrID, kstID
Dim MsgAntw As Integer
' Schleife über alle Datensätze in der Variablen
For i = 1 To XLSXmax
    ' SQL-String erstellen und Daten schreiben
    reID = Nz(DLookup("REID", "tblRechtseinheit", "RE = '" & MitarbeiterXLSXdata(i).re & "'"))
    StammNrID = Nz(DLookup("MitID", "tblStammnummer", "StammNr = " & MitarbeiterXLSXdata(i).personalnr))
    kstID = Nz(DLookup("KstID", "tblKostenstelle", "Kst = '" & MitarbeiterXLSXdata(i).kstelle & "'"))
    
    abtID = Nz(DMax("AbtID", "tblAbteilung", "kuerzel= '" & AbtXLSXdata(i).kuerzel & "'"))
    kreisID = Nz(DMax("MitKreisID", "tblMitKreis", "MitKreis= '" & MitKreisXLSXdata(i).mitKreis & "'"))
    grpID = Nz(DMax("FzgGrpID", "tblFzgGrp", "FzgGrp= '" & FzgGrpXLSXdata(i).fzgGrp & "'"))
    mitID = Nz(DMax("MitID", "tblMitarbeiter", "DKXKennung= '" & MitarbeiterXLSXdata(i).dKXKennung & "'"))
    
    sSQLAbt = "INSERT INTO tblAbteilung (kuerzel ) VALUES ('" & AbtXLSXdata(i).kuerzel & "');"
    
    sSQLKreis = "INSERT INTO tblMitKreis (MitKreis ) VALUES ('" & MitKreisXLSXdata(i).mitKreis & "');"
    
    sSQLGrp = "INSERT INTO tblFzgGrp (FzgGrp) VALUES ('" & FzgGrpXLSXdata(i).fzgGrp & "');"
    
    sSQLMit = "INSERT INTO tblMitarbeiter (Nachname,Vorname, DKX-Kennung,Email-Adresse, MitKreisID) " & _
    "VALUES('" & MitarbeiterXLSXdata(i).nachname & "', '" & MitarbeiterXLSXdata(i).vorname & "','" & _
    MitarbeiterXLSXdata(i).dKXKennung & "','" & MitarbeiterXLSXdata(i).email & "','" & MitarbeiterXLSXdata(i).mitKreis & "');"
    
    sSQLKst = "INSERT INTO tblKostenstelle (Kostenstelle, AbtID) " & _
    "VALUES ('" & MitarbeiterXLSXdata(i).kstelle & "', '" & AbtXLSXdata(i).kuerzel & "');"
    
    sSQLStNr = "INSERT INTO tblStammnummer (StammNr, ReID, KstID) " & _
    "VALUES ('" & MitarbeiterXLSXdata(i).stammNr & "', '" & MitarbeiterXLSXdata(i).re & "', '" & _
    MitarbeiterXLSXdata(i).kstelle & "');"
    
    
    Debug.Print sSQLAbt
    Debug.Print sSQLKreis
    Debug.Print sSQLGrp
    
    DoCmd.SetWarnings False


'1.######################################## tblMitKreis #################################################
    If kreisID = 0 Then
        DoCmd.RunSQL sSQLKreis
        kreisID = Nz(DMax("MitKreisID", "tblMitKreis", "MitKreis= '" & MitKreisXLSXdata(i).mitKreis & "'"))

    End If

'2.######################################## tblfzgGrp #################################################
    If grpID = 0 Then
        DoCmd.RunSQL sSQLGrp
        grpID = Nz(DMax("FzgGrpID", "tblFzgGrp", "FzgGrp= '" & FzgGrpXLSXdata(i).fzgGrp & "'"))
        
    End If

'3.######################################## tblMitarbeiter #################################################
    If mitID = 0 Then
        DoCmd.RunSQL sSQLMit
        grpID = Nz(DMax("FzgGrpID", "tblFzgGrp", "FzgGrp= '" & FzgGrpXLSXdata(i).fzgGrp & "'"))
        
    End If

'4.######################################## tblVerkMitar #################################################
     If mitID = 0 Then
        DoCmd.RunSQL sSQLStNr
        grpID = Nz(DMax("FzgGrpID", "tblFzgGrp", "FzgGrp= '" & FzgGrpXLSXdata(i).fzgGrp & "'"))
        
    End If

'5.######################################## tblAbteilung #################################################
    If abtID = 0 Then
        DoCmd.RunSQL sSQLAbt
    End If


'6######################################## tblKostenstelle #################################################

'7######################################## tblStammnummer #################################################
DoCmd.SetWarnings True
Next i


End Sub

Private Sub test()
Call ImportDaten
Call WriteXLSXDaten
End Sub


















































































































































































































