Public Function LLPrintListe5640(frm As Form, _
                             LL1 As ListLabel.ListLabel, _
                             BelegID As Long, _
                             Mode As Integer, _
                             Optional tmp As Boolean, _
                             Optional Save As Boolean, _
                             Optional SammelDruck As Boolean) As Long

    On Error GoTo Fehler

    'Mode = 2 -> Vorschau
    'Mode = 1 -> Druck
    'Mode = 4 -> Ablage
    'Mode = 0 -> Druck-Wiederholung (Druckstatus darf nicht = 0 gesetzt werden.)
    'Tmp = True -> Der Schalter wird nur bei der Vorschau noch nicht gedruckter Belege gesetzt.
    'Die Rechnungsdaten wurden zuvor in Tmp Tabellen gespeichert.
    'Save = True -> LL-Datei wird gesichert.
    'Die Option wird genutzt um die Belege zu archivieren. (Im 2-ten Durchlauf nachdem die Belege gedruckt wurden.)
    'SammelDruck -> Wird von SP52850.LLPrintSammel verwaltet.
  
    Dim SteuerText         As Variant                                       'HW 05.08.2011 Ver.: 6.1.105

    Dim Formular           As String

    Dim i                  As Integer

    Dim Msg                As Boolean

    Dim rs                 As ADODB.Recordset

    Dim rsH                As ADODB.Recordset

    Dim RS1                As ADODB.Recordset

    Dim rec1100Texte       As ADODB.Recordset                               'HW 09.07.2012

    Dim ZwSumme            As Double

    Dim DruckerDialog      As Boolean
  
    Dim SteuerPfl          As Double

    Dim SteuerFr           As Double

    Dim Ust                As Double

    Dim Betrag             As Double

    Dim SteuerPflWrg       As Double

    Dim SteuerFrWrg        As Double

    Dim UStWrg             As Double

    Dim BetragWrg          As Double

    Dim Kurs               As Double
  
    Dim sql                As String

    Dim BelegArt           As Integer

    Dim belegDatum         As Variant

    Dim BelegNr            As Long

    Dim Waehrung           As String

    Dim Skonto             As Single

    Dim SkontoTage         As Integer

    Dim nettoTage          As Integer

    Dim MwSt               As Single
  
    Dim Seite              As Long

    Dim TmpZusatz          As String

    Dim barcodeDaten       As BarcodeData

    Dim ArchivierungsModus As Integer                                       'HW 23.09.2015

    Dim idCollection       As Collection

    Dim PercentPosition    As Integer
        
    Dim strStCodeH         As String                                        'DF 29.07.2024 , Ver.: 6.7.101 : St.Code des Hauptsatzes (E-Rechnung)
        
    Dim intSteuerTextLkz   As Integer                                       'DF 29.07.2024 , Ver.: 6.7.101 : Lkz des SteuerTextes für die ganze Rechnung , wird anhand des Steuer-Schlüssel der Rechnung ermittelt.
        
    If tmp Then
        TmpZusatz = "Tmp"
    End If

    Set rs = New ADODB.Recordset
    Set rsH = New ADODB.Recordset
    Set RS1 = New ADODB.Recordset
    Set rec1100Texte = New ADODB.Recordset

    OPEN_gConn

    'CSBmk <HAUPT-RECORDSET VARIABLEN>
    rsH.Open "SELECT * FROM [5640_Haupt" & TmpZusatz & "] WHERE BelegID = " & BelegID, gConn, adOpenKeyset, adLockOptimistic

    If rsH.RecordCount > 0 Then

        'Wenn in der Hauptmaske BelegDatum/BelegNr uebergeben wurden, muessen diese im DruckOptionen Fenster angezeigt werden
        '--> in der Hauptmaske kann man ja jetzt garkein BelegDatum und -Nr uebergeben
        '70        If Not IsNull(rsH!BelegDatum) Then objDruckOptionen. = rsH!BelegDatum

        'HW 02.12.2015 Ver.: 6.4.113 AUSKOMMENTIERT! UND NEUE LOGIK UNTEN EINGEPFLEGT
        '*
        '80        If Not IsNull(rsH!BelegNr) ThenCurrentBelegDatum
        '90          If rsH!BelegNr > 0 Then objDruckOptionen.CurrentBelegNr = rsH!BelegNr
        '100       End If
        '70        If Mode = 0 Then              'Nur beim Druck von noch nicht gedruckten Belegen
        '80            If objDruckOptionen.CurrentBelegNr <> "" Then
        '    '260         LL1.LlDefineVariableExt "Kd_BelegNr", CInt(objDruckOptionen.currentBelegNr), LL_NUMERIC
        '    '290         LL1.LlDefineVariableExt "Kd_BelegNr", objDruckOptionen.currentBelegNr, LL_TEXT
        '90              rsH.Edit
        '100             rsH!BelegNr = objDruckOptionen.CurrentBelegNr
        '110             rsH.Update
        '120           End If
        '130       End If
        '*

        'HW 02.12.2015 Ver.: 6.4.113 NEUE LOGIK EINGEPFLEGT
        '*
        If Mode = 2 And tmp Then

            rsH!BelegNr = 0

        Else
            
            '<Added by: DFiebach at: 02.11.2020, Ver.: 6.6.103 >
            If val(objDruckOptionen.CurrentBelegNr) > 0 And rsH!BelegNr <> objDruckOptionen.CurrentBelegNr Then

                rsH!BelegNr = objDruckOptionen.CurrentBelegNr

            End If

            '</Added by: DFiebach at: 02.11.2020, Ver.: 6.6.103 >
                
        End If

        rsH.Update
        '*

        llCurrentFormNr = CInt(rsH.Fields("Art").value) + 37                'DH, 27.02.2017, 6.4.125, FormularNr fuer diesen Druck merken
 
        LL1.LlDefineVariableStart                                           'Variablenpuffer löschen.
        LL1.LlDefineFieldStart                                              'Variablenpuffer löschen.
    
        'DH, 06.02.2014, 6.2.102, Das Recordset aktualisieren, bevor alle Variablen im Formular gesetzt werden
        'DH, 06.03.2013, BelegDatum aus den DruckOptionen uebernehmen (sofern eingestellt)
        If objDruckOptionen.CurrentBelegDatum <> "" Then
            rsH.Fields("BelegDatum").value = objDruckOptionen.CurrentBelegDatum
            rsH.Update
        End If
    
        'HW 09.03.2015 Ver.: 6.4.104 Das ValutaDatum wird von nun an mit gespeichert!
        '******************************
        If objDruckOptionen.CurrentValutaDatum <> "" Then
            rsH.Fields("ValutaDatum").value = objDruckOptionen.CurrentValutaDatum
            rsH.Update
        End If

        '******************************

        If Mode = 3 Then
            ArchivierungsModus = 1
        Else
            ArchivierungsModus = 0
        End If

        Call LL18GestaltungFormular(LL1, rsH!Art + 37, rsH!MCode, MandantArr(1), , , ArchivierungsModus) 'HW 30.03.2012 Ver.: 6.1.111
        Call LLDefineVariablen(LL1, rsH, "Kd_")

        '<Modified by: IL at 7.26.2024, Ver.: 6.7.101 >
        '# ersetzen den erforderlichen Steuersatz im Datensatz und machen dann 0 zurück
        If rsH!MwSt = 0 And vValue <> 0 Then

            rsH!MwSt = dblUstSatz
            rsH.Update

            Call LLDefineFelder(LL1, rsH, "Kd_")                                                                      'Deklarationen

            rsH!MwSt = 0
            rsH.Update
           
        Else
            Call LLDefineFelder(LL1, rsH, "Kd_")

        End If

        '</Modified by: IL at 7.26.2024, Ver.: 6.7.101 >

        'HW 29.04.2016
        'CSBmk <OPT:BEARBEITER DRUCKEN>
        If BearbeiterDrucken Then
            LL1.LlDefineFieldExt "Bearbeiter_Drucken", "TRUE", LL_BOOLEAN
            LL1.LlDefineVariableExt "Bearbeiter_Drucken", "TRUE", LL_BOOLEAN
        Else
            LL1.LlDefineFieldExt "Bearbeiter_Drucken", "FALSE", LL_BOOLEAN
            LL1.LlDefineVariableExt "Bearbeiter_Drucken", "FALSE", LL_BOOLEAN
        End If

        '*
    
        'DF 30.03.2016
        If frm.frmParent.Check1(3).value Then
            LL1.LlDefineVariableExt "printPeriod", "TRUE", LL_BOOLEAN
        Else
            LL1.LlDefineVariableExt "printPeriod", "FALSE", LL_BOOLEAN
        End If

        '*
    
        LLAdressLand LL1, rsH!Lkz, "Kd_Land"
        LLAdressLand LL1, rsH!Lkz, "Kd_Land1"
        LL1.LlDefineVariableExt "Kd_VonDatum", "" & rsH!vonDatum, LL_TEXT
        LL1.LlDefineVariableExt "Kd_BisDatum", "" & rsH!bisDatum, LL_TEXT
        LL1.LlDefineVariableExt "ERechnungArt", 0, LL_NUMERIC               'DF 04.11.2024 , Ver.: 6.7.101
            
        LL1.LlDefineFieldExt "Anlage", "", LL_TEXT
        LL1.LlDefineVariableExt "Anlage", "", LL_TEXT

        If objDruckOptionen.CurrentBelegDatum <> "" Then                   'DH, 06.03.2013, BelegDatum aus den DruckOptionen uebernehmen (sofern eingestellt)

            '240         LL1.LlDefineVariableExt "Kd_BelegDatum", objDruckOptionen.CurrentBelegDatum, LL_DATE_LOCALIZED
            '
            '            rsH.Edit
            '            rsH.Fields("BelegDatum").Value = objDruckOptionen.CurrentBelegDatum
            '            rsH.Update
                
        Else

            If Mode = 2 Then                                                'Wenn die Vorschau aufgerufen wurde
                LL1.LlDefineVariableExt "Kd_BelegDatum", 0, LL_DATE
            End If
                
        End If

        LL1.LlDefineFieldExt "ProbeDruckText", GetZusatzText("ZusatzTexte_55710", 4), LL_TEXT 'HW 17.10.2013
    
        If Mode = 2 And Save = False Then                                   'HW 22.04.2015 Wenn ProbeDruck und die Variable "Save" auf Falsch ist!

            LL1.LlDefineFieldExt "ProbeDruck", 1, LL_NUMERIC                'HW 17.10.2013

        Else
            
            LL1.LlDefineFieldExt "ProbeDruck", 0, LL_NUMERIC                'HW 17.10.2013

        End If
    
        BelegArt = rsH!Art
        belegDatum = rsH!belegDatum
        Waehrung = rsH!Wrg1
        Skonto = rsH!ZSkto
        SkontoTage = rsH!ZSktoTage
        nettoTage = rsH!ZTage
        MwSt = rsH!MwSt
        Kurs = rsH!Kurs
        BelegNr = rsH!BelegNr
            
        'CSBmk <BARCODE>
        barcodeDaten.Seperator = ";"
        barcodeDaten.BelegNr = rsH!BelegNr
        barcodeDaten.belegDatum = IIf(IsNull(rsH!belegDatum), "", rsH!belegDatum)

        barcodeDaten.Name1 = "" & rsH.Fields("Name1").value
        barcodeDaten.Name2 = "" & rsH.Fields("Name2").value
        barcodeDaten.Adresse = "" & rsH.Fields("Straße").value
        barcodeDaten.Lkz = "" & rsH.Fields("Lkz").value
        barcodeDaten.Plz = "" & rsH.Fields("Plz").value
        barcodeDaten.Ort = "" & rsH.Fields("Ort").value
        barcodeDaten.ORTSTEIL = rsH.Fields("Ortsteil").value

        Call LL18DefineBarcode(LL1, barcodeDaten, rsH!Art + 37, "" & rsH.Fields("MCode").value)    'DH, 23.10.2013, 6.2.100, Fuer Gutschrift/Rechnung muss jeweils eine andere Formularnummer abgefragt werden
            
        '<Added by: DFiebach at: 26.07.2024, Ver.: 6.7.101 >
        'St.Code des Hauptsatzes ahnahd des gewählten St.Schl
                
        'CSBmk <KUNDEN E-RECHNUNG EINSTELLUNG>
            
        'gEnmKudnenERechnungType = modERechnung.GetKundenERechnungType(rsH.Fields("MCode").value) 'DF 23.07.2024 , Ver.: 6.7.101
            
        If IsEBelegDoc Then LL1.LlDefineVariableExt "ERechnungArt", CInt(gEnmKudnenERechnungType), LL_NUMERIC
            
        Select Case Mode

            Case 1, 4

                intSteuerTextLkz = objDruckOptionen.CurrentSteuerValue

                'CSBmk <STEUER-CODE HAUPT>
                strStCodeH = GetStCodeFromSteuerText(CStr(intSteuerTextLkz), "Rng")

        End Select

        '</Added by: DFiebach at: 26.07.2024, Ver.: 6.7.101 >
            
        'CSBmk <FOLGE-RECORDSET VARIABLEN>
        rs.Open "SELECT * FROM [5640_Folge" & TmpZusatz & "] WHERE BelegID = " & BelegID & " ORDER BY Nr", gConn, adOpenStatic, adLockReadOnly

        If rs.RecordCount > 0 Then

            Seite = 1

            Call LLDefineVariablen(LL1, rs, "Re_")

            LL1.LlDefineVariableExt "Seite", Seite, LL_NUMERIC
            LL1.LlDefineFieldExt "LetzteSeite", 0, LL_NUMERIC
            LL1.LlDefineFieldExt "Re_ZwSumme", 0, LL_NUMERIC
            LL1.LlDefineFieldExt "ZahlungsZiel", "", LL_TEXT
            LL1.LlDefineFieldExt "zZielLinks", "", LL_TEXT
            LL1.LlDefineFieldExt "zZielRechts", "", LL_TEXT
     
            'Betrag für ZahlungsZiel
            EndBetraege "5640_Folge" & TmpZusatz, BelegID, SteuerPfl, SteuerFr

            '<Removed by: DFiebach at: 19.09.2019, Ver.: 6.5.114 >
            '500             Ust = Runden((SteuerPfl * MwSt / 100), 2)
            '</Removed by: DFiebach at: 19.09.2019, Ver.: 6.5.114 >
                
            If MwSt = 0 Then                                               'IL 25.07.2024,    Nehmen den Steuerprozentsatz aus der Datenbank, wenn er im Forma 0 ist

                MwSt = GetWaehrung(rsH!WrgSchl, False).MwSt

            End If

            '<Added by: DFiebach at: 18.09.2019, Ver.: 6.5.114 >
            Ust = SteuerBetrag(CCur(SteuerPfl), Format(CCur(MwSt), "#.00"))
            '</Added by: DFiebach at: 18.09.2019, Ver.: 6.5.114 >
                
            'CSBmk <BERECHNUNG>
            Betrag = SteuerPfl + Ust + SteuerFr
            SteuerPflWrg = Runden(SteuerPfl * Kurs, 2)
            SteuerFrWrg = Runden(SteuerFr * Kurs, 2)
            UStWrg = Runden(Ust * Kurs, 2)
            BetragWrg = SteuerPflWrg + UStWrg + SteuerFrWrg

            LL1.LlDefineFieldExt "Re_SummeSteuerPfl", SteuerPfl, LL_NUMERIC
            LL1.LlDefineFieldExt "Re_SummeSteuerFr", SteuerFr, LL_NUMERIC
            LL1.LlDefineFieldExt "Re_USt", Ust, LL_NUMERIC
            LL1.LlDefineFieldExt "Re_Betrag", Betrag, LL_NUMERIC

            'HW 05.08.2011 Ver.: 6.1.105
            frm.fpSpread1(2).GetText 4, 4, SteuerText
            LL1.LlDefineFieldExt "Re_SteuerText", SteuerText, LL_TEXT

            LL1.LlDefineFieldExt "Re_SummeSteuerPflWrg", SteuerPflWrg, LL_NUMERIC
            LL1.LlDefineFieldExt "Re_SummeSteuerFrWrg", SteuerFrWrg, LL_NUMERIC
            LL1.LlDefineFieldExt "Re_UStWrg", UStWrg, LL_NUMERIC
            LL1.LlDefineFieldExt "Re_BetragWrg", BetragWrg, LL_NUMERIC

            'Soll Spalte Rabatt sichtbar sein?
            sql = "SELECT Max([Rabatt]) AS MaxRabatt "
            sql = sql & "FROM [5640_Folge" & TmpZusatz & "] WHERE BelegID = " & BelegID

            RS1.Open sql, gConn, adOpenStatic, adLockReadOnly

            LL1.LlDefineFieldExt "RabattVisible", RS1!MaxRabatt, LL_NUMERIC
            LL1.LlDefineFieldExt "KostenstellenDruck", Abs(GbKostenstellenPflicht), LL_NUMERIC 'MW 28.12.07

            'CSBmk <STEUER-TEXTE>
            objERechnung.colSteuerTexte.Clear                               'DF 30.08.2024 , Ver.: 6.7.101

            rec1100Texte.Open "SELECT * FROM [1100_Texte] WHERE Textart = 'Rng' AND Sort <= 7 Order By Lkz", gConn, adOpenStatic, adLockReadOnly 'HW 09.07.2012

            If rec1100Texte.RecordCount > 0 Then

                Do While Not rec1100Texte.EOF

                    LL1.LlDefineFieldExt "Steuertext" & CStr(rec1100Texte!Sort), "" & rec1100Texte!text, LL_TEXT
                        
                    If Not objERechnung.colSteuerTexte.ContainsKey("") Then
                        
                        Call objERechnung.colSteuerTexte.Add("" & rec1100Texte!text, CStr(rec1100Texte!Sort)) 'DF 30.08.2024 , Ver.: 6.7.101, E-Rechnung

                    End If
                        
                    rec1100Texte.MoveNext

                Loop

            Else
                
                LL1.LlDefineFieldExt "Steuertext", "", LL_TEXT

            End If
                
            rec1100Texte.Close
            Set rec1100Texte = Nothing

            LL1.LlDefineFieldExt "Steuertext", "" & gstrSteuerText, LL_TEXT 'HW 05.07.2012
            LL1.LlDefineFieldExt "SteuerSchl", intSteuerTyp, LL_NUMERIC

            'DH, 06.03.2013, Versuch die neuen Druck-Optionen einzubauen
            If objDruckOptionen.CurrentSteuertext <> "Automatisch" Then     'Wenn der Steuertext auf automatisch steht, soll der Text nicht weiter geaendert werden

                LL1.LlDefineFieldExt "Steuertext", "" & objDruckOptionen.CurrentSteuertext, LL_TEXT

                gstrSteuerText = objDruckOptionen.CurrentSteuertext         'DF 30.08.2024 , Ver.: 6.7.101, E-Rechnung
                            
            End If
                
            objERechnung.SteuerText = GetSteuerText(intSteuerTyp, SteuerFr, gstrSteuerText, objERechnung.colSteuerTexte.GetItem("2"), objERechnung.colSteuerTexte.GetItem("4"), objERechnung.colSteuerTexte.GetItem("6")) 'DF 30.08.2024 , Ver.: 6.7.101, E-Rechnung
                
        Else
            
            Msg = True

        End If

    Else
        
        Msg = True

    End If
  
    If Msg = False Then

        Formular = FormularPfad("SP56400.lst")

        'Logik aus 55710 um Belege zu archivieren -> Schleife 2 mal: 1 Vorschau mit LL_PRINT_STORAGE (Datei ins Archiv kopieren), 2 Drucken.
        'glRet = LL1.LlPrintWithBoxStart(LL_PROJECT_LIST, Formular, LL_PRINT_STORAGE, LL_BOXTYPE_BRIDGEMETER, frm.hwnd, "printing list")
        'ArbeitsplatzPfad
        If Mode < 2 Then                                                    'Druck
                
            'DRUCK
            glRet = LL1.LlPrintWithBoxStart(LL_PROJECT_LIST, Formular, LL_PRINT_NORMAL, LL_BOXTYPE_BRIDGEMETER, frm.hwnd, "printing list")

        Else                                                                'Vorschau
            
            'VORSCHAU
                
            If Save Then

                glRet = LL1.LlPrintStart(LL_PROJECT_LIST, Formular, LL_PRINT_PREVIEW)

            Else
                
                glRet = LL1.LlPrintWithBoxStart(LL_PROJECT_LIST, Formular, LL_PRINT_PREVIEW, LL_BOXTYPE_BRIDGEMETER, frm.hwnd, "printing list to preview")

            End If
                
        End If
    
        If glRet < 0 Then

            GoTo Fehler

        End If
            
        'CSBmk <ANZAHL DER KOPIEN>
        If Not Save And Mode < 2 Then

            glRet = LL1.LlPrintSetOption(LL_PRNOPT_COPIES, CLng(GetSetting("SP50000", "SP56000", "SP56430_PRNOPT_COPIES", "1")))

        Else
            
            glRet = LL1.LlPrintSetOption(LL_PRNOPT_COPIES, 1)

        End If
    
        If SammelDruck Then
            DruckerDialog = CBool(GetSetting("SP50000", "SP56000", "SP56450DruckerDialog", "-1"))
        Else
            DruckerDialog = CBool(GetSetting("SP50000", "SP56000", "SP56430DruckerDialog", "-1"))
        End If

        Call LL18PositionierungFormular(LL1, rsH!Art + 37)                  'DH, 01.07.2013, 6.2.100, Bevor gedruckt wird alle Objekte positionieren

        If objDruckOptionen.CurrentKurzrechnung Then                        'DH, 11.07.2013, 6.2.100, Wenn so eingestellt, auf den Folgeseiten nicht den gesamten Kopf drucken

            Call LL18ShortHeader(LL1)

        End If
    
        If Mode = 1 Then                                                    'Nur beim richtigen Druck und nicht bei der Druckwiederholung die Kopiensteuerung aktivieren

            Select Case programmNr

                Case 64                                                     'Lademittel Rechnung
                    Call LL18SetCopies(LL1, 1, 37, "SP56400")
    
                Case 65                                                     'Lademittel Gutschrift
                    Call LL18SetCopies(LL1, 1, 38, "SP56400")

            End Select

        End If

        'CSBmk <DRUCKAUSWAHL-DIALOG>
            
        If DruckerDialog = True Then

            If Not Save Then

                If Mode <> 4 Then glRet = LL1.LlPrintOptionsDialog(frm.hwnd, "Drucker")

                'Abbruch
                If glRet = LL_ERR_USER_ABORTED Then

                    LL1.LlPrintEnd 0                    'DH, 11.07.2013, 6.2.100, Bei Abbruch den Druckjob beenden
                    LLPrintListe5640 = LL_ERR_USER_ABORTED

                    If Not SammelDruck Then

                        If Mode = 1 Or Mode = 4 Then
                            rsH!Druck = 0
                            rsH!belegDatum = Null
                            rsH.Update
                        End If

                    End If

                    Exit Function

                End If

                If Mode < 2 Then SaveSetting "SP50000", "SP56000", "SP56430_PRNOPT_COPIES", LL1.LlPrintGetOption(LL_PRNOPT_COPIES)

            End If

        End If

        'Nach Combit ist es unbedingt notwendig, die von LlPrintSetOption gesetzte Kopienanzahl
        'durch den Aufruf von LL_PRNOPT_COPIES_SUPPORTED zu bestätigen.
        glRet = LL1.LlPrintGetOption(LL_PRNOPT_COPIES_SUPPORTED)

        Screen.MousePointer = 11

        'CSBmk <VARIABLEN DRUCKEN>
        glRet = LL1.LLPrint

        While Not rs.EOF                                                    'Solange das Ende der Posten-Tabelle nicht erreicht ist...

            DoEvents

            'Prozentbalken setzen
            PercentPosition = 100 * rs.AbsolutePosition / rs.RecordCount
            glRet = LL1.LlPrintSetBoxText("Drucken", PercentPosition)

            'Datensatzfelder der Liste bekanntmachen.
            If Trim(rs!Einheit) = "%" Then
                ZwSumme = ZwSumme + Runden((rs!Menge / 100 * rs!EPreis - rs!Menge * rs!EPreis * rs!Rabatt / 100), 2)
            Else
                ZwSumme = ZwSumme + Runden((rs!Menge * rs!EPreis - rs!Menge * rs!EPreis * rs!Rabatt / 100), 2)
            End If

            LL1.LlDefineFieldExt "Re_ZwSumme", ZwSumme, LL_NUMERIC

            Call LLDefineFelder(LL1, rs, "Re_")

            'Seitenumbruch
            If rs!SatzTyp = "S" Then

                Seite = Seite + 1
                LL1.LlDefineVariableExt "Seite", Seite, LL_NUMERIC
                glRet = LL1.LLPrint

            End If

            'Felder drucken und wenn Seitenumbruch erfolgt ist,
            'Variablen und Felder erneut drucken
            While LL1.LlPrintFields = LL_WRN_REPEAT_DATA

                Seite = Seite + 1
                LL1.LlDefineVariableExt "Seite", Seite, LL_NUMERIC
                glRet = LL1.LLPrint

            Wend

            If rs!SatzTyp = "Z" Then ZwSumme = 0

            rs.MoveNext

        Wend

        LL1.LlDefineFieldExt "LetzteSeite", 1, LL_NUMERIC
    
        'CSBmk <ZAHLUNGS-KONDITIONEN>
            
        If BelegArt = 0 Then 'Nur bei Rechnungen

            LL1.LlDefineFieldExt "zZielLinks", zZielLinks, LL_TEXT
            LL1.LlDefineFieldExt "zZielRechts", zZielRechts, LL_TEXT
                
            If objERechnung Is Nothing Then Set objERechnung = New clsERechnung      'DF 28.08.2024 , Ver.: 6.7.101
            objERechnung.ZHinweisNetto = zZielRechts
            objERechnung.ZHinweisBrutto = zZielLinks
                
        End If
            
        'CSBmk <STEUER-CODE SPEICHERN HAUPT UND FOLGE>
            
        '<Added by: DFiebach at: 11.07.2024, Ver.: 6.7.101 >
        '# Texte im Hauptsatz speichern. Damit in der PDF und ERechnung das gleiche steht.
        If Mode = 1 Or Mode = 4 Then                                                               'Nur beim Druck.

            rsH!ERechnungArt = modERechnung.GetERechnungTypeValueForDB(gEnmKudnenERechnungType)    'DF 23.07.2024 , Ver.: 6.7.101
            rsH!StCode = strStCodeH

            rsH.Update

            Call SetStCode(E_DATATYPE.Lademittelfaktura_Rechnung, 1, rsH!BelegID, intSteuerTextLkz, intSteuerTyp, tmp, GintBelegArt) ' An der Stelle wird zw. SF-RNG und -GUT nicht unterschieden, da beide in der gelichen Tabelle gespeichert werden.

        End If

        '</Added by: DFiebach at: 11.07.2024, Ver.: 6.7.101 >
            
        'Tabellen-Ausdruck beenden
        glRet = LL1.LlPrintFieldsEnd()                                      'Tabellenfelder drucken 'DH, 04.07.2013, 6.2.100, Aus der Schleife herausgenommen, damit die Seitenzaehlung korrekt ist
    
        Do

            If glRet = LL_WRN_REPEAT_DATA Then                              'Wenn auf der Seite kein Platz mehr ist

                Seite = Seite + 1                                           'Seitenanzahl erhoehen und an Formular uebergeben
                LL1.LlDefineVariableExt "Seite", Seite, LL_NUMERIC
                LL1.LlDefineFieldExt "Seite", Seite, LL_NUMERIC
                LL1.LLPrint                                                 'Auf neue Seite wechseln
                glRet = LL1.LlPrintFieldsEnd()                              'Wieder Tabellenfelder ausgeben

            End If

        Loop Until glRet <> LL_WRN_REPEAT_DATA                              'Fortfahren, solange die Meldung kommt, dass kein Platz mehr auf der Seite ist

        'DH, 03.07.2013, 6.2.100, Direkt von HW aus Sonderfaktura uebernommen
        '##################################################

        'CSBmk <ANLAGETEXT>
            
        LL1.LlPrintResetProjectState                                                                              'Druck Zurücksetzen damit Lastpage und andere diverse Funktionen für LL18 gültig sind!

        Call LL18GestaltungFormular(LL1, rsH!Art + 37, "" & rsH.Fields("MCode").value, MandantArr(1), , , ArchivierungsModus)          'Formular aufbauen

        Call LLDefineVariablen(LL1, rsH, "Kd_")                                                                   'Deklarationen

        '<Modified by: IL at 7.26.2024, Ver.: 6.7.101 >
        '# ersetzen den erforderlichen Steuersatz im Datensatz und machen dann 0 zurück
        If rsH!MwSt = 0 And vValue <> 0 Then

            rsH!MwSt = dblUstSatz

            rsH.Update

            Call LLDefineFelder(LL1, rsH, "Kd_")                                                                      'Deklarationen

            rsH!MwSt = 0

            rsH.Update
           
        Else
            
            Call LLDefineFelder(LL1, rsH, "Kd_")

        End If

        '</Modified by: IL at 7.26.2024, Ver.: 6.7.101 >
            
        'Deklarationen

        If SPLL8.bAnlageAktiv And Trim(SPLL8.strAnlageText) <> "" Then                                            'Wenn ein Anlagetext eingestellt ist!

            Debug.Print "Aktuelle Anlagen-Seite: " & Seite
    
            LL1.LlDefineFieldExt "Anlage", SPLL8.strAnlageText, LL_TEXT
            LL1.LlDefineVariableExt "Anlage", SPLL8.strAnlageText, LL_TEXT

            Seite = Seite + 1
            LL1.LlDefineVariableExt "Seite", Seite, LL_NUMERIC
            LL1.LlDefineFieldExt "Seite", Seite, LL_NUMERIC
            glRet = LL1.LLPrint

            While LL1.LlPrintFields = LL_WRN_REPEAT_DATA

                Seite = Seite + 1
                LL1.LlDefineVariableExt "Seite", Seite, LL_NUMERIC
                LL1.LlDefineFieldExt "Seite", Seite, LL_NUMERIC
                'glRet = LL1.LlPrintFields
                glRet = LL1.LLPrint

            Wend

            LL1.LlDefineFieldExt "LetzteSeite", Seite, LL_NUMERIC
        End If

        '##################################################

        'CSBmk <DRUCK BEENDEN>
        glRet = LL1.LlPrintEnd(0)
    
        'DH, 22.06.2017, ,6.4.126, Nur wenn es sich um einen scharfen Druck handelt
        If (Mode = 1 Or Mode = 4) And Not tmp Then

            'DH, 02.11.2015, 6.4.109, Aus der Speichern-Methode des Hauptforms hier her verlegt, damit das Druckkennzeichen erst gesetzt wird, wenn der Druck auch wirklich erfolgte
            If rsH.Fields("ZwAblage") = 0 Then

                rsH!Druck = 1

                If Not IsDate(rsH!belegDatum) Then rsH!belegDatum = GdtDatum

                rsH.Update

            End If
                
        End If
    
        If (Mode = 1 Or Mode = 4) And Not tmp Then

            Protokoll iAppend, "Einzeldruck -> BelegID: " & BelegID & " SteuerPfl: " & SteuerPfl & " SteuerFr: " & SteuerFr & " Ust: " & Ust & " Betrag: " & Betrag

        End If
            
        'CSBmk <ÜBERGABE AN RAB>
        'CSBmk <PDF-ARCHIVIERUNG>
        'Beim Preview-Druck Preview anzeigen und dann Preview-Datei (.LL) löschen
    
        If Mode > 1 And Mode <> 4 Then 'PrintMode = LL_PRINT_PREVIEW       auser Ablage

            If Save Then                                                                            'Archivieren

                If BelegArt = 0 Then                                                                'Rechnung

                    If rsH!ZwAblage = 0 Then
                            
                        'RECHNUNG
                            
                        'HW 19.11.2015 Hier her verschoben verschoben!
                        Call BelegAnAusgangsbuch(BelegID, E_DATATYPE.Lademittelfaktura_Rechnung)    'Belegdaten in die Ausgangsbuch Tabellen schreiben
                        Call BelegAnOpUebergeben(BelegID)                   'GW 12.3.2021, Ver.: 6.6.110
                        Call Archivieren(LL1, "LMR", BelegNr, rsH, rs)

                        Uebergabe rsH, BelegArt, "SP56430"                  'HW 02.12.2015 Hier her verschoben

                    End If

                Else                                                                                'Gutschrift

                    If rsH!ZwAblage = 0 Then
                            
                        'GUTSCHRIFT
                            
                        'HW 19.11.2015 Hier her verschoben verschoben!
                        Call BelegAnAusgangsbuch(BelegID, E_DATATYPE.Lademittelfaktura_Gutschrift)  'Belegdaten in die Ausgangsbuch Tabellen schreiben
                        Call BelegAnOpUebergeben(BelegID)                   'GW 12.3.2021, Ver.: 6.6.110
                        Call Archivieren(LL1, "LMG", BelegNr, rsH, rs)

                        Uebergabe rsH, BelegArt, "SP56430"                  'HW 02.12.2015 Hier her verschoben
  
                    End If
                End If

                Dim currentDocType As E_DATATYPE
  
                If BelegArt = 0 Then
                    currentDocType = E_DATATYPE.Lademittelfaktura_Rechnung
                Else
                    currentDocType = E_DATATYPE.Lademittelfaktura_Gutschrift
                End If
                    
                'CSBmk <EMAIL-VERSAND>
                '<Modified by: GW at 11.03.2020, Ver.: GOBD >
                If emailActivated(rsH.Fields("MCode").value, CInt(currentDocType)) Then           'DH, 21.12.2015, 6.4.115, Wenn der eMail-Versand aktiviert ist (Mandanten-/Kundenstamm)

                    If objEmailSending Is Nothing Then Set objEmailSending = New clsEmailSending
                    Set idCollection = New Collection
                    idCollection.Add BelegID

                    Call objEmailSending.StartEmailSending(frm.frmParent.cReSize.CurrScaleFactorHeight, frm.frmParent.cReSize.CurrScaleFactorWidth, voll_automatik, currentDocType, idCollection)

                End If

                '</Modified by: GW at 11.03.2020, Ver.: GOBD >
  
            Else                                                                                'Nicht archivieren
                
                'CSBmk <VORSCHAU ANZEIGEN>
                glRet = LL1.LlPreviewDisplay(Environ("Temp") & "\SP56400.LL", "", frm.hwnd)   'DH, 03.12.2014, 6.4.101, In v6.3.101 wurde der Temp-Pfad fuer die Druckvorschau auf das locale Temp-Verzeichnis gesetzt. Das entsprechende Verzeichnis muss hier natuerlich abgegriffen werden.

            End If
                
            'CSBmk <TEMP DATEI LÖSCHEN>
            glRet = LL1.LlPreviewDeleteFiles(Environ("Temp") & "\SP56400.LL", "")

        End If
    
        rs.MoveFirst

        Screen.MousePointer = 0

    End If

    If Not rsH Is Nothing Then

        If rsH.state = adStateOpen Then
            rsH.Close
        End If

        Set rsH = Nothing

    End If

    If Mode = 0 Or Mode = 1 Or Mode = 4 Then                                 'DH, 14.03.2013, Nach dem Druck/Druckwiederholung muessen BelegNr und -Datum gesperrt

        objDruckOptionen.EnableBelegDatum = False
        objDruckOptionen.EnableBelegNr = False

        printDone = True

    End If

    Exit Function
  
Exit_PrintListe:

    'HW 31.07.2013
    '##############################
    On Error Resume Next
  
    rs.Close

    If Err.number <> 0 Then Err.Clear
    Set rs = Nothing
  
    rsH.Close

    If Err.number <> 0 Then Err.Clear
    Set rsH = Nothing
  
    RS1.Close

    If Err.number <> 0 Then Err.Clear
    Set RS1 = Nothing
  
    rec1100Texte.Close

    If Err.number <> 0 Then Err.Clear
    Set rec1100Texte = Nothing
    '##############################
  
Fehler:
   
    LLPrintListe5640 = Err.number

    If Mode = 1 Or Mode = 4 Then
        If Not SammelDruck Then
            If Not rsH Is Nothing Then
                If rsH.RecordCount > 0 Then
                    rsH!Druck = 0
                    rsH!belegDatum = Null
                    rsH.Update
                End If
            End If
        End If
    End If

    Debug.Print "LL Fehler: " & glRet

    Call FehlerErklärung("SP56000B", "LLPrintListe5640()")
    GoSub Exit_PrintListe

End Function
