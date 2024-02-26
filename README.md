# VBA-Flowchart
I have too make a Flowchart for my codes.


My CODE:

Option Explicit 'Option Explicit nicht zwingend notwendig, da die Funktion in einem andere Modul aufgerufen wird

'Festlegung von öffentlichen Variablen für Weiterverwendung im nächsten Modul
Public M As Double
Public Festigkeiten As String
Public wy As Double
Public material As String
Public Länge As Double
Public Profil As String
Public Lastfall As Variant
Public foundwy As Double
Public fyd As Double

Public Function BerechneQuerschnittsdaten(Länge As Double, material As String, Profil As String, Festigkeiten As String, Lastfall As Variant, wy As Double) As Double 'Public Function für Tabellenauswertung und Rückgabe in Tabellen im Arbeitsblatt "Ausgabe"
    
    Dim ws As Worksheet 'Deklaration, dass die Variable ws für Arbeitsblätter verwendet wird
    Dim wyBereich As Range '
    Dim spalte As Range
    
    If material <> "Stahl" And material <> "Holz" Then ' Überprüfung auf gültige Materialangaben
        MsgBox "Ungültiges Material. Bitte 'Stahl' oder 'Holz' eingeben.", vbExclamation
        Exit Function
    End If

    If material = "Stahl" Then
        If Festigkeiten <> "S235" And Festigkeiten <> "S355" Then ' Überprüfung auf gültige Stahlgüte
            MsgBox "Ungültige Stahlgüte. Bitte 'S235' oder 'S355' eingeben.", vbExclamation ' Wenn keine Festigkeit eingegeben wurde wird eine Fehlermeldung ausgegeben
            Exit Function
        End If

        If Profil = "IPE" Then
            Set ws = ThisWorkbook.Sheets("Profiltafel IPE") ' Namen des IPE Stahltabellenblatts
            Set wyBereich = ws.Range("J5:J22") 'Definition der Spaltenlänge wy
           
        ElseIf Profil = "HEB" Then
            Set ws = ThisWorkbook.Sheets("Profiltafel HEB") ' Namen des HEB Stahltabellenblatts
            Set wyBereich = ws.Range("J5:J22") 'Definition der Spaltenlänge wy
           
        End If
    ElseIf material = "Holz" Then
        'Gebrauchstauglichkeit
        Dim vorhandenIyh As Double
        Dim vorhandenwyh As Double
        Dim herf As Double
        Dim Af As Double
            Select Case Profil
                Case "GL 24h"
                    herf = Round((6 * wy / 24) ^ (1 / 2), 0)
                    vorhandenwyh = (24 * (herf) ^ 2) / 6 ' gewählt wy
                    vorhandenIyh = (24 * (herf) ^ 2) / 12 ' gewählt Iy
                    Af = 24 * herf 'Fläche
                    MsgBox "Gesamter Holzquerschnitt gewählt: 24\" & herf
                    MsgBox "wy vorhanden: " & vorhandenwyh
                    MsgBox "Iy vorhanden: " & vorhandenIyh
                    MsgBox "Fläche: " & Af
                Case "GL 28h"
                    herf = Round((6 * wy / 28) ^ (1 / 2), 0)
                    vorhandenwyh = (28 * (herf) ^ 2) / 6 ' gewählt wy
                    vorhandenIyh = (28 * (herf) ^ 2) / 12 ' gewählt Iy
                    Af = 28 * herf 'Fläche
                    MsgBox "Gesamter Holzquerschnitt gewählt: 28\" & herf
                    MsgBox "wy vorhanden: " & vorhandenwyh
                    MsgBox "Iy vorhanden: " & vorhandenIyh
                    MsgBox "Fläche: " & Af
            End Select
        Debug.Print "Test: " & wy
        'Biegenormalspannung
        Dim sigmaEdh As Double
        sigmaEdh = Round(Auflagerkräfte.M * 100 / vorhandenwyh, 2)
        
            If sigmaEdh < fmgd Then
            MsgBox "Biegenormalspannung erfüllt: " & sigmaEdh & "<" & fmgd
            ElseIf sigmaEdh > fmgd Then
            MsgBox "Biegenormalspannung nicht erfüllt: " & sigmaEdh & ">" & fmgd
            End If
            
        'Schubspannungsnachweis
        Dim tauEdh As Double
        tauEdh = Round((3 / 2) * (V / Af), 2)
        
            If tauEdh < fvgd Then
            MsgBox "Schubspannungsnachweis erfüllt: " & tauEdh & "<" & fvgd
            ElseIf tauEdh > fvgd Then
            MsgBox "Schubspannungsnachweis nicht erfüllt: " & tauEdh & ">" & fvgd
            End If
        
        'Gebrauchstauglichkeitsnachweis (Überprüfung ob wmax l/300 nicht überschritten wird)
        Dim wmaxh1 As Double
        Dim wmaxh2 As Double
        Dim EHolz As Double
        
            If LF1 Then
                wmaxh1 = Round((5 * LF1 / 100 * (Länge * 100) ^ 4) / (4838400 * vorhandenIyh), 2)
                If wmaxh1 > Länge * 100 / 300 Then
                    MsgBox "Gebrauchstauglichkeit nicht erfüllt: " & wmaxh1 & ">" & Round(Länge * 100 / 300, 2)
                    MsgBox "Größerer Querschnitt erforderlich!"
                    vorhandenIy = vorhandenIy * (wmaxh1 / (Länge * 100 / 300))
                ElseIf wmaxh1 < Länge * 100 / 300 Then
                    MsgBox ("Gebrauchstauglichkeit erfüllt: " & wmaxh1 & "<" & Round(Länge * 100 / 300, 2))
                End If
            ElseIf LF2 Or LF3 Then
                wmaxh2 = Round((LF2 * (Länge * 100) ^ 3) / (604800 * vorhandenIyh), 2)
                If wmaxh2 > Länge * 100 / 300 Then
                    MsgBox "Gebrauchstauglichkeit nicht erfüllt: " & wmaxh2 & ">" & Round(Länge * 100 / 300, 2)
                    MsgBox "Größerer Querschnitt wird berechnet"
                    vorhandenIy = vorhandenIy * (wmaxh2 / (Länge * 100 / 300))
                ElseIf wmaxh2 < Länge * 100 / 300 Then
                    MsgBox ("Gebrauchstauglichkeit erfüllt: " & wmaxh2 & "<" & Round(Länge * 100 / 300, 2))
                End If
            End If

'Wertrückgabe in das Arbeitsblatt Ausgabe für das ausgewählte Holzprofil

Dim NachweiserfülltHolz As Range
ThisWorkbook.Sheets("Ausgabe").Range("K5").Value = herf
    Select Case Profil
        Case "GL 24h"
        ThisWorkbook.Sheets("Ausgabe").Range("L5").Value = 24
        Case "GL 28h"
        ThisWorkbook.Sheets("Ausgabe").Range("L5").Value = 28
    End Select
ThisWorkbook.Sheets("Ausgabe").Range("M5").Value = vorhandenwyh
ThisWorkbook.Sheets("Ausgabe").Range("N5").Value = vorhandenIyh
ThisWorkbook.Sheets("Ausgabe").Range("O5").Value = Af

'Einfärben einer Zelle wenn bedingung erfüllt
Set NachweiserfülltHolz = ThisWorkbook.Sheets("Ausgabe").Range("P5")
' Überprüfung der Erfüllung des Nachweises
    If sigmaEdh < fmgd And tauEdh < fvgd And (wmaxh1 < Länge * 100 / 300 Or wmaxh2 < Länge * 100 / 300) Then
        ' Wenn die Bedingung erfüllt ist, färben Sie die Zelle grün ein
        NachweiserfülltHolz.Interior.Color = RGB(0, 255, 0)  ' RGB-Wert für Grün
    Else
        ' Andernfalls, färben Sie die Zelle rot ein
        NachweiserfülltHolz.Interior.Color = RGB(255, 0, 0)  ' RGB-Wert für Rot
    End If

Exit Function
    End If

    If Len(Profil) = 0 Then ' Überprüfung auf gültiges Profil
        MsgBox "Bitte ein Profil angeben.", vbExclamation
        Exit Function
    End If

 For Each spalte In wyBereich
        If spalte.Value >= wy Then
            ' Wenn der Wert größer oder gleich dem gesuchten wy-Wert ist: speichern und Schleife beenden
            foundwy = spalte.Value
            Exit For
        End If
    Next spalte
    
    ' Befehl: Wenn der gesuchte wy-Wert nicht genau gefunden wurde, nimm den nächsthöheren Wert
    If foundwy = 0 Then
        foundwy = Application.WorksheetFunction.Min(wyBereich)
    End If
    
    
        'Iy: Zelle rechts neben foundWy
        Dim foundIyCell As Range
        Set foundIyCell = spalte.Offset(0, 1)

        ' Lies den Wert aus der Zelle
        Dim foundIy As Double
        foundIy = foundIyCell.Value
        
        If foundIy = 0 Then
        MsgBox "Ungültiger Wert für foundIy. Bitte überprüfen Sie die Eingabedaten.", vbExclamation
        Exit Function
        End If
        
        ' s
        Dim foundValueCells As Range
        Set foundValueCells = spalte.Offset(0, -6)

        ' s
        Dim foundValues As Double
        foundValues = foundValueCells.Value
        
        ' t
        Dim foundValueCellt As Range
        Set foundValueCellt = spalte.Offset(0, -5)

        ' t
        Dim foundValuet As Double
        foundValuet = foundValueCellt.Value
        
        ' Sy
        Dim foundSyCell As Range
        Set foundSyCell = spalte.Offset(0, 5)
        
        ' Sy
        Dim foundSy As Double
        foundSy = foundSyCell.Value
    
        'Profil
        Dim foundProfilCell As Range
        Set foundProfilCell = spalte.Offset(0, -9)
        
        'Profil
        Dim foundProfil As Double
        foundProfil = foundProfilCell.Value
    
 'Überprüfung ob wmax l/300 nicht überschritten wird (Gebrauchstauglichkeit)
    Dim wmax1 As Double
    Dim wmax2 As Double
    Dim EStahl As Double
    EStahl = 21000
    
  
  Do
    If LF1 Then
        wmax1 = Round((LF1 / 100) * ((Länge * 100) ^ 4) / (8064000 * foundIy), 2)
        Debug.Print wmax1
        If wmax1 > Länge * 100 / 300 Then
            'MsgBox "Gebrauchstauglichkeit nicht erfüllt: " & wmax1 & ">" & Round(Länge * 100 / 300, 2)
            MsgBox "Größerer Querschnitt wird berechnet"
            foundIy = foundIy * (wmax1 / (Länge * 100 / 300))
        ElseIf wmax1 < Länge * 100 / 300 Then
            'MsgBox ("Gebrauchstauglichkeit erfüllt: " & wmax1 & "<" & Round(Länge * 100 / 300, 2))
        End If
    ElseIf LF2 Or LF3 Then
        wmax2 = Round((LF2 / 100 * (Länge * 100) ^ 3) / (1008000 * foundIy), 2)
        If wmax2 > Länge * 100 / 300 Then
            'MsgBox "Gebrauchstauglichkeit nicht erfüllt: " & wmax2 & ">" & Round(Länge * 100 / 300, 2)
            MsgBox "Größerer Querschnitt wird berechnet"
            
            foundIy = foundIy * (wmax2 / (Länge * 100 / 300))
        ElseIf wmax2 < Länge * 100 / 300 Then
            'MsgBox ("Gebrauchstauglichkeit erfüllt: " & wmax2 & "<" & Round(Länge * 100 / 300, 2))
        End If
    End If
    
    ' Wiederholen Sie die Suche nach Iy und andere Berechnungen für das neue Profil
    Dim foundIyCellNew As Range
    Set foundIyCellNew = foundIyCell.Offset(1, 0)
    foundIy = foundIyCellNew.Value
    ' andere Werte: wy, s, t
    
    'wy: Zelle links neben foundWy
        Dim foundwyCell As Range
        Set foundwyCell = spalte.Offset(1, 0)

        ' Lies den Wert aus der Zelle
        foundwy = foundwyCell.Value
        
        ' s
        'Dim foundValueCells As Range
        Set foundValueCells = spalte.Offset(1, -6)

        ' s
        'Dim foundValues As Double
        foundValues = foundValueCells.Value
        
        ' t
        'Dim foundValueCellt As Range
        Set foundValueCellt = spalte.Offset(1, -5)

        ' t
        'Dim foundValuet As Double
        foundValuet = foundValueCellt.Value
        
        'Sy
        Set foundSyCell = spalte.Offset(1, 5)
        
        'Sy
        foundSy = foundSyCell.Value
        
        'Profil
        Set foundProfilCell = spalte.Offset(1, -9)
        
        'Profil
        foundProfil = foundProfilCell.Value
    

Loop While wmax1 > Länge * 100 / 300 Or wmax2 > Länge * 100 / 300
        
    
    ' gefundener wy-Wert
    MsgBox "Gefundenes wy für Material " & material & ": " & foundwy & " cm" & ChrW(179) 'Der ASCII-Code für "hoch 3" ist 179

    
    ' gefundener Iy-Wert
    MsgBox "Gefundenes Iy für Material " & material & ": " & foundIy & " cm^4" 'hoch 4
    
    ' gefundenes s
    MsgBox "Gefundenes s für Material " & material & ": " & foundValues & " mm" ' in mm
    
    ' gefundenes t
    MsgBox "Gefundenes t für Material " & material & ": " & foundValuet & " mm" ' in mm
    
    ' gefundenes Sy
    MsgBox "Gefundenes Sy für Material " & material & ": " & foundSy & " cm" & ChrW(179) 'Der ASCII-Code für "hoch 3" ist 179
    
    
    Debug.Print "Gesuchtes wy: " & wy 'Ausgabe wenn wy gefunden
     
    If foundwy <> 0 Then
    Debug.Print "Gefundenes wy: " & foundwy
    Debug.Print "Gefundenes Iy: " & foundIy
    Else
    Debug.Print "Wert nicht gefunden." 'Ausgabe wenn wy nicht gefunden
    End If
     
   
'Biegenormalspannungsnachweis
    sigma = Auflagerkräfte.M * 100 / foundwy

    If Festigkeiten = "S235" Then
        If sigma < 23.5 Then
        MsgBox ("Biegenormalspannungsnachweis erfüllt: " & Round(sigma, 2) & "< 23.5")
        ElseIf sigma > 23.5 Then
        MsgBox ("Biegenormalspannungsnachweis nicht erfüllt: " & Round(sigma, 2) & "> 23.5")
        End If
    ElseIf Festigkeiten = "S355" Then
        If sigma < 35.5 Then
        MsgBox ("Biegenormalspannungsnachweis erfüllt: " & Round(sigma, 2) & "< 35.5")
        ElseIf sigma > 35.5 Then
        MsgBox ("Biegenormalspannungsnachweis nicht erfüllt: " & Round(sigma, 2) & "> 35.5")
        End If
    End If
    

'Schubspannung
Dim tauEds As Double

tauEds = Round((V * foundSy) / (foundIy * (foundValues / 10)), 2)
            If tauEds < BWyMitM.fyd / (3 ^ 0.5) Then
            MsgBox "Schubspannungsnachweis erfüllt: " & tauEds & "<" & Round(BWyMitM.fyd / (3 ^ 0.5), 2)
            ElseIf tauEds > BWyMitM.fyd / (3 ^ 0.5) Then
            MsgBox "Schubspannungsnachweis nicht erfüllt: " & tauEds & ">" & Round(BWyMitM.fyd / (3 ^ 0.5), 2)
            End If

'Gebrauchstauglichkeit
If LF1 Then
        wmax1 = Round((LF1 / 100) * ((Länge * 100) ^ 4) / (8064000 * foundIy), 2)
        Debug.Print wmax1
        If wmax1 > Länge * 100 / 300 Then
            MsgBox "Gebrauchstauglichkeit nicht erfüllt: " & wmax1 & ">" & Round(Länge * 100 / 300, 2)
            MsgBox "Größerer Querschnitt wird berechnet"
            foundIy = foundIy * (wmax1 / (Länge * 100 / 300))
        ElseIf wmax1 < Länge * 100 / 300 Then
            MsgBox ("Gebrauchstauglichkeit erfüllt: " & wmax1 & "<" & Round(Länge * 100 / 300, 2))
        End If
    ElseIf LF2 Or LF3 Then
        wmax2 = Round((LF2 / 100 * (Länge * 100) ^ 3) / (1008000 * foundIy), 2)
        If wmax2 > Länge * 100 / 300 Then
            MsgBox "Gebrauchstauglichkeit nicht erfüllt: " & wmax2 & ">" & Round(Länge * 100 / 300, 2)
            MsgBox "Größerer Querschnitt wird berechnet"
            foundIy = foundIy * (wmax2 / (Länge * 100 / 300))
        ElseIf wmax2 < Länge * 100 / 300 Then
            MsgBox ("Gebrauchstauglichkeit erfüllt: " & wmax2 & "<" & Round(Länge * 100 / 300, 2))
        End If
    End If
    
'Wertrückgabe in das Arbeitsblatt Ausgabe für das ausgewählte Stahlprofil
Dim ausgabews As Worksheet
Dim Nachweiserfüllt As Range
ThisWorkbook.Sheets("Ausgabe").Range("B5").Value = Profil & " " & foundProfil
ThisWorkbook.Sheets("Ausgabe").Range("C5").Value = foundwy
ThisWorkbook.Sheets("Ausgabe").Range("D5").Value = foundIy
ThisWorkbook.Sheets("Ausgabe").Range("E5").Value = foundValues
ThisWorkbook.Sheets("Ausgabe").Range("F5").Value = foundValuet
ThisWorkbook.Sheets("Ausgabe").Range("G5").Value = foundSy

'Einfärben einer Zelle wenn bedingung erfüllt
Set Nachweiserfüllt = ThisWorkbook.Sheets("Ausgabe").Range("H5")
' Überprüfung der Erfüllung des Nachweises
    If sigma < 23.5 Or sigma < 35.5 And tauEds < BWyMitM.fyd / (3 ^ 0.5) And (wmax1 < Länge * 100 / 300 Or wmax2 < Länge * 100 / 300) Then
        ' Wenn die Bedingung erfüllt ist, färben Sie die Zelle grün ein
        Nachweiserfüllt.Interior.Color = RGB(0, 255, 0)  ' RGB-Wert für Grün
    Else
        ' Andernfalls, färben Sie die Zelle rot ein
        Nachweiseerfüllt.Interior.Color = RGB(255, 0, 0)  ' RGB-Wert für Rot
    End If

End Function
'Public Function SchnellesM() As Double
'SchnellesM = Auflagerkräfte.M
'End Function
