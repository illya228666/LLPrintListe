Public Function LLPrintListe(frm As Form, _
                             LL1 As ListLabel.ListLabel, _
                             BelegID As Long, _
                             Mode As Integer, _
                             Optional tmp As Boolean, _
                             Optional Save As Boolean, _
                             Optional SammelDruck As Boolean) As Long
    'Mode = 4 -> Ablage
    'Mode = 3 -> Archivierung
    'Mode = 2 -> Vorschau
    'Mode = 1 -> Druck
    'Mode = 0 -> Druck-Wiederholung (Druckstatus darf nicht = 0 gesetzt werden.)
    'Tmp = True -> Der Schalter wird nur bei der Vorschau noch nicht gedruckter Belege gesetzt.
    'Die Rechnungsdaten wurden zuvor in Tmp Tabellen gespeichert.
    'Save = True -> LL-Datei wird gesichert.
    'Die Option wird genutzt um die Belege zu archivieren. (Im 2-ten Durchlauf nachdem die Belege gedruckt wurden.)
    'SammelDruck -> Wird von SP52850.LLPrintSammel verwaltet.
  
    On Error GoTo Fehler

    Dim Formular                 As String

    Dim i                        As Integer

    'Dim j                        As Integer

    Dim Msg                      As Boolean

    Dim rs                       As ADODB.Recordset

    Dim rsH                      As ADODB.Recordset

    Dim RS1                      As ADODB.Recordset

    Dim ZwSumme                  As Double

    Dim DruckerDialog            As Boolean
  
    Dim SteuerPfl                As Double

    Dim SteuerFr                 As Double

    Dim Ust                      As Double

    Dim UstTMP                   As Double

    Dim Betrag                   As Double

    Dim SteuerPflWrg             As Double

    Dim SteuerFrWrg              As Double

    Dim UStWrg                   As Double

    Dim BetragWrg                As Double

    Dim Kurs                     As Double
  
    Dim sql                      As String

    Dim BelegArt                 As Integer

    Dim belegDatum               As Variant

    Dim BelegNr                  As Long

    Dim Waehrung                 As String

    Dim Skonto                   As Single

    Dim SkontoTage               As Integer

    Dim nettoTage                As Integer

    Dim MwSt                     As Single

    Dim Seite                    As Long

    Dim TmpZusatz                As String
  
    Dim barcodeDaten             As BarcodeData
        
    Dim SteuerText               As Variant                                'HW 05.11.2010 Ver.: 5.4.122 - Ver.: 6.1.101
        
    Dim strSteuerTextAusFormular As String                                 'DF 11.07.2024 , Ver.: 6.7.101 : der entgültige SteuerText wird im Formular ausgewählt. Daher wird es zum Speichern daraus geholt.
        
    Dim strZahlungsTextNetto     As String                                 'DF 11.07.2024 , Ver.: 6.7.101
         
    Dim strZahlungsText          As String                                 'DF 11.07.2024 , Ver.: 6.7.101

    Dim rec1100Texte             As ADODB.Recordset                   'HW 09.07.2012 Ver.: 6.1.114
  
    Dim ArchivierungsModus       As Integer                                'HW 23.09.2015

    Dim idCollection             As Collection

    Dim PercentPosition          As Integer
        
    Dim strLogBelegArt           As String                                 'DF 16.01.2019 , Ver.: 6.5.109 : DruckArt -String für LOG-Datei
        
    Dim lngBelegNrKres           As Integer                                'DF 05.02.2019 , Ver.: 6.5.109 : Nummer des NrKreises
        
    Dim strStCodeH               As String                                 'DF 29.07.2024 , Ver.: 6.7.101 : St.Code des Hauptsatzes (E-Rechnung)
                
    Dim intSteuerTextLkz         As Integer                                'DF 29.07.2024 , Ver.: 6.7.101 : Lkz des SteuerTextes für die ganze Rechnung , wird anhand des Steuer-Schlüssel der Rechnung ermittelt.
        
    'CSBmk <LOG START>
    Select Case Mode
        
        Case 0         'DRUCK WIEDERHOLUNG
                 
            Protokoll iAppend, ">DRUCK START (MODUS: DRUCK WIEDERHOLUNG) -> BelegID: " & BelegID & ""
                 
        Case 1         'DRUCK
            
            Protokoll iAppend, ">DRUCK START (MODUS: DRUCK) -> BelegID: " & BelegID & ""
            
        Case 2         'VORSCHAU
            
            Protokoll iAppend, ">DRUCK START (MODUS: VORSCHAU) -> BelegID: " & BelegID & ""
            
        Case 3         'ARCHIVIERUNG
        
            Protokoll iAppend, ">DRUCK START (MODUS: ARCHIVIERUNG) -> BelegID: " & BelegID & ""

        Case 4         'ABLAGE
        
            Protokoll iAppend, ">DRUCK START (MODUS: ABLAGE) -> BelegID: " & BelegID & ""
            
    End Select
        
    If tmp Then

        TmpZusatz = "Tmp"

    End If

    Set rs = New ADODB.Recordset
    Set rsH = New ADODB.Recordset
    Set RS1 = New ADODB.Recordset
    Set rec1100Texte = New ADODB.Recordset

    OPEN_gConn
        
    'CSBmk <HAUPT-RECORDSET VARIABLEN>
    rsH.Open "SELECT * FROM [2800_Haupt" & TmpZusatz & "] WHERE BelegID = " & BelegID, gConn, adOpenKeyset, adLockOptimistic
 
    If rsH.RecordCount > 0 Then

        LL1.LlDefineVariableStart                                           'Variablenpuffer löschen.

        LL1.LlDefineFieldStart                                              'Variablenpuffer löschen.
            
        j = 1

        If Mode = 3 Then
            ArchivierungsModus = 1
        Else
            ArchivierungsModus = 0
        End If

        llCurrentFormNr = CInt(objDruckOptionen.FormularNr)                'IL 25.10.2024 , Ver.: 6.7.101 :  rsH!art + 35 ------> objDruckOptionen.FormularNr

        Call LL18GestaltungFormular(LL1, objDruckOptionen.FormularNr, "" & rsH.Fields("MCode").value, MandantArr(1), , , ArchivierungsModus)   'IL 25.10.2024 , Ver.: 6.7.101 :  rsH!art + 35 ------> objDruckOptionen.FormularNr
        j = 2

        Call LLDefineVariablen(LL1, rsH, "Kd_")
        j = 3
            
        Call LLDefineTexte(LL1)                                             'DF 24.10.2024 , Ver.: 6.7.101 : ZusatzTexte usw.
            
        '<Modified by: IL at 7.26.2024, Ver.: 6.7.101 >
        '# ersetzen den erforderlichen Steuersatz im Datensatz und machen dann 0 zurück
        If rsH!MwSt = 0 And CInt(vValue) <> 0 Then

            rsH!MwSt = dblUstSatz

            rsH.Update

            Call LLDefineFelder(LL1, rsH, "Kd_")                            'Deklarationen

            rsH!MwSt = 0

            rsH.Update

        Else
            
            Call LLDefineFelder(LL1, rsH, "Kd_")

        End If

        '</Modified by: IL at 7.26.2024, Ver.: 6.7.101 >
        j = 4
            
        'CSBmk <OPT:BEARBEITER DRUCKEN>
        If BearbeiterDrucken Then

            LL1.LlDefineVariableExt "Bearbeiter_Drucken", "TRUE", LL_BOOLEAN
            LL1.LlDefineFieldExt "Bearbeiter_Drucken", "TRUE", LL_BOOLEAN

        Else
            
            LL1.LlDefineVariableExt "Bearbeiter_Drucken", "FALSE", LL_BOOLEAN
            LL1.LlDefineFieldExt "Bearbeiter_Drucken", "FALSE", LL_BOOLEAN

        End If
            
        'CSBmk <OPT:ADRESSE AUF FOLGESEITEN DRUCKEN>
        If blnFolgeseitenKurzDrucken Then                                   'Added by: GW at: 24.04.2019, Ver.: 6.5.111

            LL1.LlDefineVariableExt "FolgeseitenKurz_Drucken", "TRUE", LL_BOOLEAN
            LL1.LlDefineFieldExt "FolgeseitenKurz_Drucken", "TRUE", LL_BOOLEAN

        Else
            
            LL1.LlDefineVariableExt "FolgeseitenKurz_Drucken", "FALSE", LL_BOOLEAN
            LL1.LlDefineFieldExt "FolgeseitenKurz_Drucken", "FALSE", LL_BOOLEAN

        End If
            
        'CSBmk <OPT:BRUTTO-NETTO UMRECHNUNG>
        If GesamtIstBrutto Then                                             'HW 24.05.2016 Ver.: 6.4.120
            LL1.LlDefineFieldExt "GesamtIstBrutto", "TRUE", LL_BOOLEAN
        Else
            LL1.LlDefineFieldExt "GesamtIstBrutto", "FALSE", LL_BOOLEAN
        End If
            
        If objDruckOptionen.CurrentBelegDatum <> "" Then                    'DH, 11.07.2013, BelegDatum aus den DruckOptionen uebernehmen (sofern eingestellt)

            LL1.LlDefineVariableExt "Kd_BelegDatum", objDruckOptionen.CurrentBelegDatum, LL_DATE_LOCALIZED

            LL1.LlDefineFieldExt "Kd_BelegDatum", objDruckOptionen.CurrentBelegDatum, LL_DATE_LOCALIZED

        Else

            If Mode = 2 Then                                                'Wenn die Vorschau aufgerufen wurde

                LL1.LlDefineVariableExt "Kd_BelegDatum", 0, LL_DATE
                LL1.LlDefineFieldExt "Kd_BelegDatum", 0, LL_DATE

            End If
                
        End If
  
        LL1.LlDefineFieldExt "ProbeDruckText", ZusatzText(4, "55710"), LL_TEXT 'HW 16.10.2013
        LL1.LlDefineVariableExt "ERechnungArt", 0, LL_NUMERIC               'DF 04.11.2024 , Ver.: 6.7.101

        If Mode = 2 Then                                                    'HW 16.10.2013 Wenn ProbeDruck Dann
                
            LL1.LlDefineFieldExt "ProbeDruck", 1, LL_NUMERIC                'IL 21.10.2024 , Ver.: 6.7.101 :    mode ----> 1; Um die Beschriftung anzuzeigen, muss der Parameter gleich 1 und nicht 2 sein

        Else
            
            LL1.LlDefineFieldExt "ProbeDruck", 0, LL_NUMERIC                'HW 16.10.2013

        End If

        Call DefineZusatztext(rsH, LL1)                                     'MW 13.11.08 Ver.: 5.4.119 Zusatztext
            
        BelegArt = rsH!Art
        belegDatum = rsH!belegDatum
        Waehrung = rsH!Wrg1
        Skonto = rsH!ZSkto
        SkontoTage = rsH!ZSktoTage
        nettoTage = rsH!ZTage
        MwSt = rsH!MwSt
        Kurs = rsH!Kurs
        BelegNr = rsH!BelegNr

        Select Case GintBelegArt
            
            Case 0
                
                strLogBelegArt = "Rechnungsdruck"
                
            Case 1
                
                strLogBelegArt = "Gutschriftsdruck"
                
            Case 2
                
                strLogBelegArt = "Angebot"
                
            Case 3
                
                strLogBelegArt = "Auftragsbestätigung"
            
        End Select                  'Added by: DFiebach at: 16.01.2019, Ver.: 6.5.109
            
        '<Added by: DFiebach at: 26.07.2024, Ver.: 6.7.101 >
        'St.Code des Hauptsatzes ahnahd des gewählten St.Schl
               
        'CSBmk <KUNDEN E-RECHNUNG EINSTELLUNG>
        If GesamtIstBrutto Then gEnmKudnenERechnungType = eERechnungType.None                            'DF 04.09.2024 , Ver.: 6.7.101, keine ERechnung bei Brutto-Netto Umrechnung
        
        If IsEBelegDoc Then LL1.LlDefineVariableExt "ERechnungArt", CInt(gEnmKudnenERechnungType), LL_NUMERIC

        Select Case Mode
            
            Case 1, 4                                                      'IL 07.11.2024 , Ver.: 6.7.101 : Ablage hinzugefügt
                            
                intSteuerTextLkz = objDruckOptionen.CurrentSteuerValue
                
                'CSBmk <STEUER-CODE HAUPT>
                strStCodeH = GetStCodeFromSteuerText(CStr(intSteuerTextLkz), "Rng")

        End Select

        '</Added by: DFiebach at: 26.07.2024, Ver.: 6.7.101 >
            
        'CSBmk <FOLGE-RECORDSET VARIABLEN>
        rs.Open "SELECT * FROM [2800_Folge" & TmpZusatz & "] WHERE BelegID = " & BelegID & " ORDER BY Nr", gConn, adOpenStatic, adLockReadOnly

        j = 5
    
        If rs.RecordCount > 0 Then

            Seite = 1

            Call LLDefineVariablen(LL1, rs, "Re_")

            LL1.LlDefineVariableExt "Seite", Seite, LL_NUMERIC
            LL1.LlDefineFieldExt "LetzteSeite", 0, LL_NUMERIC
            LL1.LlDefineFieldExt "Re_ZwSumme", 0, LL_NUMERIC
            LL1.LlDefineFieldExt "ZahlungsZiel", "", LL_TEXT
            LL1.LlDefineFieldExt "ZahlungsZielNetto", "", LL_TEXT

            j = 6
                
            EndBetraege "2800_Folge" & TmpZusatz, BelegID, SteuerPfl, SteuerFr  'Betrag für ZahlungsZiel

            'CSBmk <STEUER>
            If MwSt = 0 Then                                               'IL 25.07.2024, Nehmen den Steuerprozentsatz aus der Datenbank, wenn er im Form 0 ist

                MwSt = GetWaehrung(rsH!WrgSchl, False).MwSt

            End If
                
            If GesamtIstBrutto Then                                         'HW 24.05.2016 Ver.: 6.4.120

                UstTMP = SteuerBetrag(CCur(SteuerPfl), Format(CCur(MwSt), "#.00"), False)

                Ust = SteuerPfl - UstTMP

            Else

                Ust = SteuerBetrag(CCur(SteuerPfl), Format(CCur(MwSt), "#.00"))
       
            End If

            If GesamtIstBrutto Then                                         'HW 24.05.2016 Ver.: 6.4.120

                SteuerPfl = (SteuerPfl - Ust)

            End If
                
            'CSBmk <BERECHNUNG>
            Betrag = SteuerPfl + Ust + SteuerFr                             'DF 19.02.2020 , Ver.: 6.5.114 : Runden -> RundenMiVz
            SteuerPflWrg = RundenMitVz(SteuerPfl * Kurs, 2)
            SteuerFrWrg = RundenMitVz(SteuerFr * Kurs, 2)
            UStWrg = Runden(Ust * Kurs, 2)
            BetragWrg = SteuerPflWrg + UStWrg + SteuerFrWrg

            LL1.LlDefineFieldExt "Re_EPreisDezStellen", postCommaPreis, LL_NUMERIC       'DH, 27.10.2017, 6.5.101, Einstellung aus den Systemparameter uebergeben
            LL1.LlDefineFieldExt "Re_SummeSteuerPfl", SteuerPfl, LL_NUMERIC
            LL1.LlDefineFieldExt "Re_SummeSteuerFr", SteuerFr, LL_NUMERIC
            LL1.LlDefineFieldExt "Re_USt", Ust, LL_NUMERIC
            LL1.LlDefineFieldExt "Re_Betrag", Betrag, LL_NUMERIC
            LL1.LlDefineFieldExt "Re_SummeSteuerPflWrg", SteuerPflWrg, LL_NUMERIC
            LL1.LlDefineFieldExt "Re_SummeSteuerFrWrg", SteuerFrWrg, LL_NUMERIC
            LL1.LlDefineFieldExt "Re_UStWrg", UStWrg, LL_NUMERIC
            LL1.LlDefineFieldExt "Re_BetragWrg", BetragWrg, LL_NUMERIC

            'CSBmk <STEUER-TEXTE>
            objERechnung.colSteuerTexte.Clear

            frm.fpSpread1(2).GetText 4, 4, SteuerText                       'HW 05.11.2010 Ver.: 5.4.122 - Ver.: 6.1.101
            LL1.LlDefineFieldExt "Re_SteuerText", SteuerText, LL_TEXT

            rec1100Texte.Open "SELECT * FROM [1100_Texte] WHERE Textart = 'Rng' AND Sort <= 7", gConn, adOpenStatic, adLockReadOnly 'HW 09.07.2012 Ver.: 6.1.114
     
            If rec1100Texte.RecordCount > 0 Then

                Do While Not rec1100Texte.EOF

                    LL1.LlDefineFieldExt "Steuertext" & rec1100Texte!Sort, "" & rec1100Texte!text, LL_TEXT
                        
                    If Not objERechnung.colSteuerTexte.ContainsKey(CStr(rec1100Texte!Sort)) Then
                        
                        Call objERechnung.colSteuerTexte.Add("" & rec1100Texte!text, CStr(rec1100Texte!Sort)) 'DF 30.08.2024 , Ver.: 6.7.101, E-Rechnung

                    End If
                        
                    rec1100Texte.MoveNext

                Loop

            Else
                
                LL1.LlDefineFieldExt "Steuertext", "", LL_TEXT

            End If
                
            rec1100Texte.Close
            Set rec1100Texte = Nothing
                
            LL1.LlDefineFieldExt "Steuertext", "" & gstrSteuerText, LL_TEXT 'HW 05.07.2012  Ver.: 6.1.129
            LL1.LlDefineFieldExt "SteuerSchl", intSteuerTyp, LL_NUMERIC
                                
            objERechnung.SteuerText = GetSteuerText(intSteuerTyp, SteuerFr, gstrSteuerText, objERechnung.colSteuerTexte.GetItem("2"), objERechnung.colSteuerTexte.GetItem("4"), objERechnung.colSteuerTexte.GetItem("6")) 'DF 30.08.2024 , Ver.: 6.7.101, E-Rechnung
                
            'CSBmk <ANLAGEN-TEXT>
            LL1.LlDefineFieldExt "AnlagenText", "", LL_TEXT                 'HW 01.07.2013
            LL1.LlDefineVariableExt "AnlagenText", "", LL_TEXT              'HW 01.07.2013
                
            'CSBmk <VON / BIS DATUM>
            If IsDate(rsH!vonDatum) Then
                LL1.LlDefineVariableExt "Kd_VonDatum", rsH!vonDatum, LL_TEXT
            Else
                LL1.LlDefineVariableExt "Kd_VonDatum", "", LL_TEXT
            End If

            If IsDate(rsH!bisDatum) Then
                LL1.LlDefineVariableExt "Kd_BisDatum", rsH!bisDatum, LL_TEXT
            Else
                LL1.LlDefineVariableExt "Kd_BisDatum", "", LL_TEXT
            End If
    
            j = 7
                
            'CSBmk <RABATT SICHTBAR?>
            sql = "SELECT Max([Rabatt]) AS MaxRabatt "
            sql = sql & "FROM [2800_Folge" & TmpZusatz & "] WHERE BelegID = " & BelegID
            RS1.Open sql, gConn, adOpenStatic, adLockReadOnly

            LL1.LlDefineFieldExt "RabattVisible", RS1!MaxRabatt, LL_NUMERIC
            RS1.Close
            j = 8

            'CSBmk <BELEG MIT LIEFERSCHIENARTIKEL?>
            sql = "SELECT TOP 1 SatzTyp "                                   'MW 26.04.05
            sql = sql & "FROM [2800_Folge" & TmpZusatz & "] WHERE SatzTyp = 'L' AND BelegID = " & BelegID
            RS1.Open sql, gConn, adOpenStatic, adLockReadOnly

            LL1.LlDefineFieldExt "LSArtikel", RS1.RecordCount, LL_NUMERIC
            RS1.Close

            LL1.LlDefineFieldExt "KostenstellenDruck", Abs(GbKostenstellenPflicht), LL_NUMERIC 'MW 28.12.07

            j = 9

        Else
            
            Msg = True

        End If

    Else
        
        Msg = True

    End If

    j = 10
  
    If Msg = False Then

        Formular = FormularPfad("SP52800.lst")

        'glRet = LL1.LlPreviewSetTempPath(ArbeitsplatzPfad & "\1\")
            
        j = 11
    
        'Logik aus 55710 um Belege zu archivieren -> Schleife 2 mal: 1 Vorschau mit LL_PRINT_STORAGE (Datei ins Archiv kopieren), 2 Drucken.
        'glRet = LL1.LlPrintWithBoxStart(LL_PROJECT_LIST, Formular, LL_PRINT_STORAGE, LL_BOXTYPE_BRIDGEMETER, frm.hwnd, "printing list")
        'ArbeitsplatzPfad
            
        If Mode < 2 Then
                
            'DRUCK
            glRet = LL1.LlPrintWithBoxStart(LL_PROJECT_LIST, Formular, LL_PRINT_NORMAL, LL_BOXTYPE_BRIDGEMETER, frm.hwnd, "Druck")

            j = 12

            '<Added by: IL at: 07.11.2024, Ver.: 6.7.101 >
            '# Druck in die LL Datei.
        ElseIf Mode = 4 Then

            'ABLAGE
            glRet = LL1.LlPrintWithBoxStart(LL_PROJECT_LIST, Formular, LL_PRINT_PREVIEW, LL_BOXTYPE_BRIDGEMETER, frm.hwnd, "Druck")

        Else
            '</Added by: IL at: 07.11.2024, Ver.: 6.7.101 >
                    
            'VORSCHAU

            If Save Then

                '895                 glRet = LL1.LlPrintStart(LL_PROJECT_LIST, Formular, LL_PRINT_PREVIEW)

                glRet = LL1.LlPrintWithBoxStart(LL_PROJECT_LIST, Formular, LL_PRINT_PREVIEW, LL_BOXTYPE_BRIDGEMETER, frm.hwnd, "Archivierung")
                    
                j = 13

            Else
                
                glRet = LL1.LlPrintWithBoxStart(LL_PROJECT_LIST, Formular, LL_PRINT_PREVIEW, LL_BOXTYPE_BRIDGEMETER, frm.hwnd, "Vorschau")

                j = 14

            End If

        End If

        If glRet < 0 Then

            j = 15

            GoTo Fehler

        End If
            
        'CSBmk <ANZAHL DER KOPIEN>
        If Not Save And Mode < 2 Then
            glRet = LL1.LlPrintSetOption(LL_PRNOPT_COPIES, CLng(GetSetting("SP50000", "SP52800", "SP52830_PRNOPT_COPIES", "1")))
            glRet = LL1.LlPrintGetOption(LL_PRNOPT_COPIES_SUPPORTED)
        Else
            glRet = LL1.LlPrintSetOption(LL_PRNOPT_COPIES, 1)
            glRet = LL1.LlPrintGetOption(LL_PRNOPT_COPIES_SUPPORTED)
        End If

        j = 16
            
        If SammelDruck Then
            DruckerDialog = CBool(GetSetting("SP50000", "SP52800", "SP52850DruckerDialog", "-1"))
        Else
            DruckerDialog = CBool(GetSetting("SP50000", "SP52800", "SP52830DruckerDialog", "-1"))
        End If
    
        Call LL18PositionierungFormular(LL1, objDruckOptionen.FormularNr)                   'IL 25.10.2024 , Ver.: 6.7.101 :  rsH!art + 35 ------> objDruckOptionen.FormularNr

        If objDruckOptionen.CurrentKurzrechnung Then                        'DH, 11.07.2013, 6.2.100, Wenn so eingestellt, auf den Folgeseiten nicht den gesamten Kopf drucken

            Call LL18ShortHeader(LL1)

        End If

        If Mode = 1 Then                                                    'Nur beim richtigen Druck und nicht bei der Wiederholung

            Call LL18SetCopies(LL1, 1, objDruckOptionen.FormularNr, "SP52800")  'IL 25.10.2024 , Ver.: 6.7.101 :  rsH!art + 35 ------> objDruckOptionen.FormularNr

        End If
            
        'CSBmk <DRUCKERAUSWAHL-DIALOG>
        If DruckerDialog = True Then                                        'Druckdialog

            If Not Save Then

                If Mode <> 4 Then glRet = LL1.LlPrintOptionsDialog(frm.hwnd, "Drucker")

                Select Case glRet                                           '<Added by: DFiebach at: 27.01.2022, Ver.: 6.6.112
                    
                    Case Is < 0
                            
                        Protokoll iAppend, "##### -> FEHLER : LL, NUMMER: " & CStr(glRet) & ", SF-" & strLogBelegArt & ", BelegNr = " & BelegNr & ", BelegID =" & BelegID & ""
                            
                        GetMsgFromLLErrorCode glRet
                            
                        LL1.LlPrintEnd 0

                        LLPrintListe = glRet
                        
                        If Not SammelDruck Then

                            If Mode = 1 Or Mode = 4 Then                   'IL 07.11.2024 , Ver.: 6.7.101 :  Ablage hinzugefügt

                                j = 17

                                rsH!Druck = 0
                                rsH!belegDatum = Null                       'DF 13.02.2019 , Ver.: 6.5.109 : BelegDatum nur wenn Beleg erfolgreich gedruckt wurde
                                rsH.Update

                                j = 18

                            End If
                            
                        End If
                            
                        glRet = LL1.LlPreviewDeleteFiles(ArbeitsplatzPfad & "\SP52800.LL", "") 'DF 08.03.2023 , Ver.: 6.6.118 : Vorschau-Datei beim Abbruch des DruckerAuswahl-Dialoges löschen.
                            
                        Exit Function
                             
                End Select

                If Mode < 2 Then SaveSetting "SP50000", "SP52800", "SP52830_PRNOPT_COPIES", LL1.LlPrintGetOption(LL_PRNOPT_COPIES)

            End If

            j = 16

        End If
    
        'Nach Combit ist es unbedingt notwendig, die von LlPrintSetOption gesetzte Kopienanzahl
        'durch den Aufruf von LL_PRNOPT_COPIES_SUPPORTED zu bestätigen.
        glRet = LL1.LlPrintGetOption(LL_PRNOPT_COPIES_SUPPORTED)
            
        '######## 1. HIER BELEG-NR ZIEHEN
            
        'CSBmk <BELEG-NR VERGABE>
        If Mode = 1 Or Mode = 4 Then                                       'IL 07.11.2024 , Ver.: 6.7.101 :  Ablage hinzugefügt

            If objDruckOptionen.CurrentBelegNr <> "" And objDruckOptionen.CurrentBelegNr <> "0" Then

                If Mode = 2 And tmp Then
                    rsH!BelegNr = 0
       
                Else
                    rsH!BelegNr = objDruckOptionen.CurrentBelegNr
      
                End If
                    
                Protokoll iAppend, "Beleg-Nummer aus DruckOptionen uebernommen -> BelegID: " & BelegID & ", BelegNr:" & rsH!BelegNr & ""
                    
                rsH.Update

            Else
                    
                '<Modified by: DFiebach at 01.04.2019, Ver.: 6.5.110 >
                ' # Überprüfung auf Vorlagen hinzugefügt
                If rsH!ZwAblage = 0 Then
                    
                    BelegNr = GetBelegNr(GintBelegNrKreisNr, BelegID, GintBelegArt, lngBelegNrKres, False, programmNr) 'DF 14.11.2024 , Ver.: 6.7.101 : GintBelegArt + 8 -> GintBelegNrKreisNr. GintBelegArt passt hier nicht mehr wegen der neune BelegArten (Ang, AufBest), da diese andrere Nr-Bereich in NrKresen haben.

                    '<Added by: DFiebach at: 01.04.2019, Ver.: 6.5.110 >
                    If BelegNr > 0 Then

                        gLngBelegNr = BelegNr                               'Added by: GW at: 03.04.2019, Ver.: 6.5.110
                 
                        rsH!BelegNrKReis = lngBelegNrKres
                        rsH!BelegNr = BelegNr
                        rsH.Update

                    Else
                    
                        rsH!Druck = 0
                        rsH!belegDatum = Null                               'DF 13.02.2019 , Ver.: 6.5.109 : BelegDatum nur wenn Beleg erfolgreich gedruckt wurde
                        rsH.Update
                    
                        Protokoll iAppend, "##### -> ABBRUCH DURCH BENUTZER, BELEG-NR NICHT FORTLAUFEND, SF-" & strLogBelegArt & ", BelegNr = " & BelegNr & ", BelegID =" & BelegID & ""
                        
                        LL1.LlPrintEnd 0

                        LLPrintListe = LL_ERR_USER_ABORTED
                            
                        glRet = LL1.LlPreviewDeleteFiles(ArbeitsplatzPfad & "\SP52800.LL", "")
                            
                        Exit Function
                        
                    End If

                    '</Added by: DFiebach at: 01.04.2019, Ver.: 6.5.110 >
                        
                Else
                       
                    BelegNr = 0
                       
                End If

                '</Modified by: DFiebach at 01.04.2019, Ver.: 6.5.110 >
                    
            End If

            Call LLDefineVariablen(LL1, rsH, "Kd_")                                                                   'Deklarationen

            '<Modified by: IL at 7.26.2024, Ver.: 6.7.101 >
            '#  ersetzen den erforderlichen Steuersatz im Datensatz und machen dann 0 zurück
            If rsH!MwSt = 0 And CInt(vValue) <> 0 Then

                rsH!MwSt = dblUstSatz

                rsH.Update

                Call LLDefineFelder(LL1, rsH, "Kd_")                                                                      'Deklarationen

                rsH!MwSt = 0

                rsH.Update
               
            Else
                
                Call LLDefineFelder(LL1, rsH, "Kd_")

            End If

            '</Modified by: IL at 7.26.2024, Ver.: 6.7.101 >                                                                 'Deklarationen

        End If

        'CSBmk <BARCODE DEFINIEREN>
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

        Call LL18DefineBarcode(LL1, barcodeDaten, objDruckOptionen.FormularNr, rsH.Fields("MCode").value) 'Barcode im Formular definieren       'IL 25.10.2024 , Ver.: 6.7.101 :  rsH!art + 35 ------> objDruckOptionen.FormularNr

        Screen.MousePointer = 11

        'CSBmk <VARIABLEN DRUCKEN>
        glRet = LL1.LLPrint
            
        Protokoll iAppend, "Beleg-Positionen werden gedruckt -> BelegID: " & BelegID & ""
            
        While Not rs.EOF                                                    'Solange das Ende der Posten-Tabelle nicht erreicht ist...

            j = 19

            DoEvents

            'Prozentbalken setzen
            PercentPosition = 100 * rs.AbsolutePosition / rs.RecordCount
            glRet = LL1.LlPrintSetBoxText("Drucken", PercentPosition)
                
            'Datensatzfelder der Liste bekanntmachen.                       'DF 19.02.2020 , Ver.: 6.5.114 : Runden -> RundenMiVz
            If Trim(rs!Einheit) = "%" Then
                    
                'ORIG
                'ZwSumme = ZwSumme + RundenMitVz((rs!Menge / 100 * rs!EPreis - rs!Menge * rs!EPreis * rs!Rabatt / 100), 2)
                    
                'DF 03.03.2025 , Ver.: 6.7.106: NEU -> rs!Menge / 100 führte zum falschen Ergebnis, -> rs!EPreis / 100 analog zum fpSread-Formel.
                ZwSumme = ZwSumme + RundenMitVz((rs!Menge * rs!EPreis / 100 - rs!Menge * rs!EPreis / 100 * rs!Rabatt / 100), 2)

            Else
                
                ZwSumme = ZwSumme + RundenMitVz((rs!Menge * rs!EPreis - rs!Menge * rs!EPreis * rs!Rabatt / 100), 2)

            End If

            LL1.LlDefineFieldExt "Re_EPreisDezStellen", postCommaPreis, LL_NUMERIC      'DH, 27.10.2017, 6.5.101, Einstellung aus den Systemparametern uebergeben

            LL1.LlDefineFieldExt "Re_ZwSumme", ZwSumme, LL_NUMERIC

            Call LLDefineFelder(LL1, rs, "Re_")

            'Seitenumbruch
            If rs!SatzTyp = "S" Then

                Seite = Seite + 1

                LL1.LlDefineVariableExt "Seite", Seite, LL_NUMERIC
                LL1.LlDefineFieldExt "Seite", Seite, LL_NUMERIC

                glRet = LL1.LLPrint

                j = 20

            End If
                
            'Felder drucken und wenn Seitenumbruch erfolgt ist,
            'Variablen und Felder erneut drucken
    
            'HW 01.07.2013 Wenn AnlageText eingestellt wurde!
            '##################################################
            '                     ANLAGENTEXT
            '##################################################
            LL1.LlDefineFieldExt "Anlage", "", LL_TEXT
            LL1.LlDefineVariableExt "Anlage", "", LL_TEXT
            '##################################################

            While LL1.LlPrintFields = LL_WRN_REPEAT_DATA

                Seite = Seite + 1

                LL1.LlDefineVariableExt "Seite", Seite, LL_NUMERIC
                LL1.LlDefineFieldExt "Seite", Seite, LL_NUMERIC

                glRet = LL1.LLPrint

            Wend

            If rs!SatzTyp = "Z" Then ZwSumme = 0
            rs.MoveNext

        Wend

        LL1.LlDefineFieldExt "LetzteSeite", 1, LL_NUMERIC
            
        'DF 04.06.2020 , Ver.: 6.6.102 : Zahlungs-Konditionen auch bei Gutschriften drucken
            
        'CSBmk <ZAHLUNGS-KONDITIONEN>
            
        '<Modified by: IL at 15.10.2024, Ver.: 6.7.101 >
        '#   BelegDatum -------> objDruckOptionen.CurrentBelegDatum.
        strZahlungsText = ZahlungsZiel(objDruckOptionen.CurrentBelegDatum, Betrag, Waehrung, Skonto, SkontoTage, nettoTage, objDruckOptionen.CurrentValutaDatum)
        strZahlungsTextNetto = ZahlungsZielNetto(objDruckOptionen.CurrentBelegDatum, Betrag, Waehrung, Skonto, SkontoTage, nettoTage, objDruckOptionen.CurrentValutaDatum)
        '</Modified by: IL at 15.10.2024, Ver.: 6.7.101 >
            
        LL1.LlDefineFieldExt "ZahlungsZiel", strZahlungsText, LL_TEXT      'DH, 16.02.2015, 6.4.103, ValutaDatum aus den DruckOptionen uebernehmen
        LL1.LlDefineFieldExt "ZahlungsZielNetto", strZahlungsTextNetto, LL_TEXT
            
        If objERechnung Is Nothing Then Set objERechnung = New clsERechnung      'DF 28.08.2024 , Ver.: 6.7.101
        objERechnung.ZHinweisNetto = strZahlungsTextNetto
        objERechnung.ZHinweisBrutto = strZahlungsText
            
        'CSBmk <STEUER-CODE SPEICHERN HAUPT UND FOLGE>
            
        '<Added by: DFiebach at: 11.07.2024, Ver.: 6.7.101 >

        If Mode = 1 Or Mode = 4 Then                                                               'Nur beim Druck.

            rsH!ERechnungArt = modERechnung.GetERechnungTypeValueForDB(gEnmKudnenERechnungType)    'DF 23.07.2024 , Ver.: 6.7.101
            rsH!StCode = strStCodeH
            
            rsH.Update

            Call SetStCode(E_DATATYPE.Sonderfaktura_Rechnung, 1, rsH!BelegID, intSteuerTextLkz, intSteuerTyp, tmp, GintBelegArt) ' An der Stelle wird zw. SF-RNG und -GUT nicht unterschieden, da beide in der gelichen Tabelle gespeichert werden.

        End If

        '</Added by: DFiebach at: 11.07.2024, Ver.: 6.7.101 >
            
        'Tabellen-Ausdruck beenden
        Do
            
            glRet = LL1.LlPrintFieldsEnd()

            If glRet = LL_WRN_REPEAT_DATA Then

                Seite = Seite + 1

                LL1.LlDefineVariableExt "Seite", Seite, LL_NUMERIC
                LL1.LlDefineFieldExt "Seite", Seite, LL_NUMERIC

                'Neue Seite
                LL1.LLPrint

                j = 21

            End If

        Loop Until glRet <> LL_WRN_REPEAT_DATA
  
        'HW 01.07.2013 Wenn AnlageText eingestellt wurde!
        '##################################################
            
        'CSBmk <ANLAGETEXT>
            
        LL1.LlPrintResetProjectState                                                                              'Druck Zurücksetzen damit Lastpage und andere diverse Funktionen für LL18 gültig sind!

        Call LL18GestaltungFormular(LL1, objDruckOptionen.FormularNr, "" & rsH.Fields("MCode").value, MandantArr(1), , , ArchivierungsModus)          'IL 25.10.2024 , Ver.: 6.7.101 :  rsH!art + 35 ------> objDruckOptionen.FormularNr

        Call LLDefineVariablen(LL1, rsH, "Kd_")                                                                   'Deklarationen

        '<Modified by: IL at 7.26.2024, Ver.: 6.7.101 >
        '# ersetzen den erforderlichen Steuersatz im Datensatz und machen dann 0 zurück
        If rsH!MwSt = 0 And CInt(vValue) <> 0 Then

            rsH!MwSt = dblUstSatz

            rsH.Update

            Call LLDefineFelder(LL1, rsH, "Kd_")                                                                  'Deklarationen

            rsH!MwSt = 0

            rsH.Update
           
        Else
            
            Call LLDefineFelder(LL1, rsH, "Kd_")

        End If

        '</Modified by: IL at 7.26.2024, Ver.: 6.7.101 >

        If SPLL8.bAnlageAktiv And Trim(SPLL8.strAnlageText) <> "" Then                                            'Wenn ein Anlagetext eingestellt ist!
    
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
             
                glRet = LL1.LLPrint

            Wend

            LL1.LlDefineFieldExt "LetzteSeite", Seite, LL_NUMERIC

        End If

        '##################################################
     
        'CSBmk <DRUCK BEENDEN>
        glRet = LL1.LlPrintEnd(0)

        j = 22
     
        If (Mode = 1 Or Mode = 4) And Not tmp Then                         'IL 07.11.2024 , Ver.: 6.7.101 : Ablage hinzugefügt

            Protokoll iAppend, ">EINZELDRUCK BEENDET -> BelegID: " & BelegID & " SteuerPfl: " & SteuerPfl & " SteuerFr: " & SteuerFr & " Ust: " & Ust & " Betrag: " & Betrag

        End If
            
        'CSBmk <ÜBERGABE AN RAB>
            
        'CSBmk <PDF-ARCHIVIERUNG>
        '
        'Beim Preview-Druck Preview anzeigen und dann Preview-Datei (.LL) löschen
        'HW 23.04.2014 Ver.: 6.2.105 3 übergeben! fürs archivieren!
        If Mode > 1 And Mode <> 4 Then 'PrintMode = LL_PRINT_PREVIEW       auser Ablage

            If Save Then

                '<Modified by: IL at 09.10.2024, Ver.: 6.7.101 >
                Select Case BelegArt

                    Case 0 'RECHNUNG
    
                        If rsH!ZwAblage = 0 Then

                            Protokoll iAppend, ">BELEG AN RAB uebergeben -> BelegID: " & BelegID & ""

                            Call BelegAnAusgangsbuch(BelegID, E_DATATYPE.Sonderfaktura_Rechnung, GesamtIstBrutto, CCur(SteuerPfl))       'Belegdaten in die Ausgangsbuch Tabellen schreiben

                            Protokoll iAppend, ">BELEG AN OP uebergeben -> BelegID: " & BelegID & ""

                            Call BelegAnOpUebergeben(BelegID)                   'GW 12.3.2021, Ver.: 6.6.110

                            Protokoll iAppend, ">BELEG AN ARCHIV uebergeben -> BelegID: " & BelegID & ""

                            Call ArchivierenPDF(LL1, "SFR", BelegNr, rsH, rs)

                            Protokoll iAppend, ">UEBERAGABEN AN RAB/OP/FIBU BEENDET -> BelegID: " & BelegID & ""

                        End If
    
                    Case 1 'GUTSCHRIFT
    
                        If rsH!ZwAblage = 0 Then

                            Protokoll iAppend, ">BELEG AN RAB uebergeben -> BelegID: " & BelegID & ""

                            Call BelegAnAusgangsbuch(BelegID, E_DATATYPE.Sonderfaktura_Gutschrift, GesamtIstBrutto, CCur(SteuerPfl))      'Belegdaten in die Ausgangsbuch Tabellen schreiben

                            Protokoll iAppend, ">BELEG AN OP uebergeben -> BelegID: " & BelegID & ""

                            Call BelegAnOpUebergeben(BelegID)                   'GW 12.3.2021, Ver.: 6.6.110

                            Protokoll iAppend, ">BELEG AN ARCHIV uebergeben -> BelegID: " & BelegID & ""

                            Call ArchivierenPDF(LL1, "SFG", BelegNr, rsH, rs)

                            Protokoll iAppend, ">UEBERAGABEN AN RAB/OP/FIBU BEENDET -> BelegID: " & BelegID & ""

                        End If
    
                    Case 2 'ANGEBOT
    
                        If rsH!ZwAblage = 0 Then

                            Protokoll iAppend, ">BELEG AN ARCHIV uebergeben -> BelegID: " & BelegID & ""

                            Call ArchivierenPDF(LL1, "SFA", BelegNr, rsH, rs)

                            Protokoll iAppend, ">UEBERAGABEN AN ARCHIV BEENDET -> BelegID: " & BelegID & ""

                        End If
    
                    Case 3 'AUFTRAGSBESTETIGUNG
                            
                        If rsH!ZwAblage = 0 Then
                            
                            Protokoll iAppend, ">BELEG AN ARCHIV uebergeben -> BelegID: " & BelegID & ""

                            Call ArchivierenPDF(LL1, "SFB", BelegNr, rsH, rs)

                            Protokoll iAppend, ">UEBERAGABEN AN ARCHIV BEENDET -> BelegID: " & BelegID & ""

                        End If

                End Select

                'Orig: 1695                If BelegArt = 0 Then
                '
                '                        'RECHNUNG
                '
                '1700                    If rsH!ZwAblage = 0 Then
                '
                '1705                        Protokoll iAppend, ">BELEG AN RAB uebergeben -> BelegID: " & BelegID & ""
                '
                '1710                        Call BelegAnAusgangsbuch(BelegID, E_DATATYPE.Sonderfaktura_Rechnung, GesamtIstBrutto, CCur(SteuerPfl))       'Belegdaten in die Ausgangsbuch Tabellen schreiben
                '
                '1715                        Protokoll iAppend, ">BELEG AN OP uebergeben -> BelegID: " & BelegID & ""
                '
                '1720                        Call BelegAnOpUebergeben(BelegID)                   'GW 12.3.2021, Ver.: 6.6.110
                '
                '1725                        Protokoll iAppend, ">BELEG AN ARCHIV uebergeben -> BelegID: " & BelegID & ""
                '
                '1730                        Call ArchivierenPDF(LL1, "SFR", BelegNr, rsH, rs)
                '
                '1735                        Protokoll iAppend, ">UEBERAGABEN AN RAB/OP/FIBU BEENDET -> BelegID: " & BelegID & ""
                '
                '                        End If
                '
                '                    Else
                '
                '                        'GUTSCHRIFT
                '
                '1740                    If rsH!ZwAblage = 0 Then
                '
                '1745                        Protokoll iAppend, ">BELEG AN RAB uebergeben -> BelegID: " & BelegID & ""
                '
                '1750                        Call BelegAnAusgangsbuch(BelegID, E_DATATYPE.Sonderfaktura_Gutschrift, GesamtIstBrutto, CCur(SteuerPfl))      'Belegdaten in die Ausgangsbuch Tabellen schreiben
                '
                '1755                        Protokoll iAppend, ">BELEG AN OP uebergeben -> BelegID: " & BelegID & ""
                '
                '1760                        Call BelegAnOpUebergeben(BelegID)                   'GW 12.3.2021, Ver.: 6.6.110
                '
                '1765                        Protokoll iAppend, ">BELEG AN ARCHIV uebergeben -> BelegID: " & BelegID & ""
                '
                '1770                        Call ArchivierenPDF(LL1, "SFG", BelegNr, rsH, rs)
                '
                '1775                        Protokoll iAppend, ">UEBERAGABEN AN RAB/OP/FIBU BEENDET -> BelegID: " & BelegID & ""
                '
                '                        End If
                '
                '                    End If

                Dim currentDocType As E_DATATYPE
                    
                Select Case BelegArt
                    
                    Case 0
                        
                        currentDocType = E_DATATYPE.Sonderfaktura_Rechnung
                        
                    Case 1
                        
                        currentDocType = E_DATATYPE.Sonderfaktura_Gutschrift
                        
                    Case 2
                        
                        currentDocType = E_DATATYPE.Sonderfaktura_Angebot
                        
                    Case 3
                        
                        currentDocType = E_DATATYPE.Sonderfaktura_Auftragsbestetigung
                    
                End Select
    
                'Orig:1830           If BelegArt = 0 Then
                '1835                    currentDocType = E_DATATYPE.Sonderfaktura_Rechnung
                '                    Else
                '1840                    currentDocType = E_DATATYPE.Sonderfaktura_Gutschrift
                '                    End If
                    
                '</Modified by: >IL at 09.10.2024, Ver.: 6.7.101 >
                    
                'CSBmk <EMAIL-VERSAND>
                If emailActivated(rsH.Fields("MCode").value, CInt(currentDocType)) Then           'DH, 21.12.2015, 6.4.114, Wenn der eMail-Versand aktiviert ist (Mandanten-/Kundenstamm)
                        
                    Protokoll iAppend, ">EMAIL VERSAND -> BelegID: " & BelegID & ""
                        
                    If objEmailSending Is Nothing Then Set objEmailSending = New clsEmailSending   'Modified by: GW at 21.02.2020, Ver.: GOBD_EMAIL
    
                    Set idCollection = New Collection
                    idCollection.Add BelegID

                    If UCase(frm.name) = "FRMSP52831" Then                                        'Einzeldruck
                        Call objEmailSending.StartEmailSending(frm.frmParent.cReSize.CurrScaleFactorHeight, frm.frmParent.cReSize.CurrScaleFactorWidth, voll_automatik, currentDocType, idCollection)
                    Else                                                                          'Sammeldruck
                        Call objEmailSending.StartEmailSending(frm.cReSize.CurrScaleFactorHeight, frm.cReSize.CurrScaleFactorWidth, voll_automatik, currentDocType, idCollection)
                    End If

                End If

            Else
                                                
                'CSBmk <VORSCHAU ANZEIGEN>
                glRet = LL1.LlPreviewDisplay(ArbeitsplatzPfad & "\SP52800.LL", "", frm.hwnd)

                j = 23

            End If
                
            'CSBmk <TEMP DATEI LÖSCHEN>
            glRet = LL1.LlPreviewDeleteFiles(ArbeitsplatzPfad & "\SP52800.LL", "")

            j = 24

        End If
    
        rs.MoveFirst

        Screen.MousePointer = 0

    End If
        
    rsH.Close

    Set rsH = Nothing

    'DH, 11.07.2013, Nach dem Druck/Druckwiederholung muessen BelegNr und -Datum gesperrt werden
    If Mode = 0 Or Mode = 1 Or Mode = 4 Then                               'IL 07.11.2024 , Ver.: 6.7.101 : Ablage hinzugefügt

        objDruckOptionen.EnableBelegDatum = False
        objDruckOptionen.EnableBelegNr = False
        objDruckOptionen.EnableValutaDatum = False                          'DH, 17.02.2015, 6.4.103, Das neue Feld Valuta Datum muss bei der Druckwiederholung auch gesperrt werden

        printDone = True

    End If
        
    'CSBmk <LOG ENDE>
    Select Case Mode
        
        Case 0         'DRUCK WIEDERHOLUNG
                 
            Protokoll iAppend, ">DRUCK ENDE (MODUS: DRUCK WIEDERHOLUNG) -> BelegID: " & BelegID & ""
                 
        Case 1         'DRUCK
            
            Protokoll iAppend, ">DRUCK ENDE (MODUS: DRUCK) -> BelegID: " & BelegID & ""
            
        Case 2         'VORSCHAU
            
            Protokoll iAppend, ">DRUCK ENDE (MODUS: VORSCHAU) -> BelegID: " & BelegID & ""
            
        Case 3         'ARCHIVIERUNG
        
            Protokoll iAppend, ">DRUCK ENDE (MODUS: ARCHIVIERUNG) -> BelegID: " & BelegID & ""

        Case 4         'ABLAGE                                             'IL 07.11.2024 , Ver.: 6.7.101 :  Ablage hinzugefügt

            Protokoll iAppend, ">DRUCK ENDE (MODUS: ABLAGE) -> BelegID: " & BelegID & ""
            
    End Select
        
    Exit Function
  
Fehler:
   
    LLPrintListe = Err.number

    If Mode = 1 Or Mode = 4 Then                                           'IL 07.11.2024 , Ver.: 6.7.101 :  Ablage hinzugefügt

        If Not SammelDruck Then

            If Not rsH Is Nothing Then

                If rsH.RecordCount > 0 Then
                    rsH!Druck = 0
                    rsH!belegDatum = Null                                   'DF 13.02.2019 , Ver.: 6.5.109 : BelegDatum nur wenn Beleg erfolgreich gedruckt wurde
                    rsH.Update
                End If

            End If

        End If
            
    End If

    If glRet <> 0 Then
        Call FehlerErklärung("SP52800B", "LLPrintListe LL-Fehler: " & glRet & ", j = " & j)
    Else
        Call FehlerErklärung("SP52800B", "LLPrintListe")
    End If

End Function
