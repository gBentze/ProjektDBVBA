# ProjektDBVBA
Option Compare Database
Option Explicit

'Tabelle Abteilung, Mitkreis, typeFzgGrp
Const dateipfad = "D:\DantenbankAccess\Abschlussphase"
Const xlBlattName As String = "Tabelle1"

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
    DKXKennung As String
    Nachname As String
    Vorname As String
    Email As String
    mitKreis As String
    lngMitKreis As Long
    re As String
    personalnr As Long
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
Dim vKuerzel As Variant
Dim vMitKreis As Variant
Dim vFzgGrp As Variant
Dim vDKXKennung As Variant
Dim vNachname As Variant
Dim vVorname As Variant
Dim vEmail As Variant
Dim vMitKreis  As Variant
Dim vRE As Variant
Dim vPersonal As Variant
Dim VKostenstelle As Variant

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
        sCol = "A"
        vPersonal = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
        sCol = "B"
        vNachname = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
        sCol = "C"
        vVorname = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
        sCol = "D"
        vDKXKennung = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
        sCol = "E"
        vEmail = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
        sCol = "F"
        vKuerzel = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
        sCol = "G"
        vRE = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
        sCol = "H"
        VKostenstelle = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
        sCol = "I"
        vMitKreis = xlsApp.Worksheets(xlBlattName).Range(sCol & iRowS & ":" & sCol & iRowL)
        sCol = "J"
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
                MitarbeiterXLSXdata(i).DKXKennung = vDKXKennung(i, 1)
                MitarbeiterXLSXdata(i).Email = vEmail(i, 1)
                MitarbeiterXLSXdata(i).Nachname = vNachname(i, 1)
                MitarbeiterXLSXdata(i).Vorname = vVorname(i, 1)
                MitarbeiterXLSXdata(i).mitKreis = vMitKreis(i, 1)
                MitarbeiterXLSXdata(i).re = vRE(i, 1)
                MitarbeiterXLSXdata(i).kstelle = VKostenstelle(i, 1)
                MitarbeiterXLSXdata(i).personalnr = vPersonal(i, 1)

            Next i
        End With

    Set xlsApp = Nothing
End Sub
    


Private Sub WriteXLSXDaten() 'bShowInfo As Boolean)
''' Daten aus Variable in Tabelle übertragen
Dim i As Integer
'Dim lngMitKreisID As Long
Dim sSQLAbt As String
Dim sSQLKreis As String
Dim sSQLGrp As String
Dim abtId As Long, grpId As Long, kreisID As Long, mitID As Long, reID As Long, persID, kstID
Dim MsgAntw As Integer

    For i = 1 To XLSXmax ' Schleife über alle Datensätze in der Variablen
        ' SQL-String erstellen und Daten schreiben

        
        reID = Nz(DLookup("REID", "tblRechtseinheit", "RE = '" & MitarbeiterXLSXdata(i).re & "'"))
        persID = Nz(DLookup("MitID", "tblStammnummer", "StammNr = " & MitarbeiterXLSXdata(i).personalnr))
        kstID = Nz(DLookup("KstID", "tblKostenstelle", "Kst = '" & MitarbeiterXLSXdata(i).kstelle & "'"))
        
        abtId = Nz(DMax("AbtID", "tblAbteilung", "kuerzel= '" & AbtXLSXdata(i).kuerzel & "'"))
        kreisID = Nz(DMax("MitKreisID", "tblMitKreis", "MitKreis= '" & MitKreisXLSXdata(i).mitKreis & "'"))
        grpId = Nz(DMax("FzgGrpID", "tblFzgGrp", "FzgGrp= '" & FzgGrpXLSXdata(i).fzgGrp & "'"))
        mitID = Nz(DMax("MitID", "tblMitarbeiter", "DKXKennung= '" & FzgGrpXLSXdata(i).fzgGrp & "'"))
        
        sSQLAbt = "INSERT INTO tblAbteilung (kuerzel ) VALUES ('" & AbtXLSXdata(i).kuerzel & "');"
        sSQLKreis = "INSERT INTO tblMitKreis (MitKreis ) VALUES ('" & MitKreisXLSXdata(i).mitKreis & "');"
        sSQLGrp = "INSERT INTO tblFzgGrp (FzgGrp) VALUES ('" & FzgGrpXLSXdata(i).fzgGrp & "');"
        
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
        
        If grpId = 0 Then
            DoCmd.RunSQL sSQLGrp
            
        End If
'3.######################################## tblMitarbeiter #################################################


'4.######################################## tblVerkMitar #################################################

'5.######################################## tblAbteilung #################################################
        If abtId = 0 Then
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



