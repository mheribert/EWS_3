Option Explicit                       ' alle Variablen muessen deklariert werden
Const max_anzahl As Double = 150
Public sklasse As TextBox
' Deklarationen
Public Type TEinzelwertung
    TP_ID       As Integer      ' ID des Tanzpaares
    StartNo     As Long         ' Startnummer des Teilnehmers
    Wertung(8)  As Long         ' die Platzwertungen
    Punkt(8)    As Double       ' die Punkte
    WR_ID(8)    As Integer      ' ID des Wertungsrichters
    Platz       As Long         ' der ermittelte Platz
    klasse      As String       ' Startklasse
End Type

Private Type TMatrixZelle
  Anzahl        As Long                    ' Anzahl der Wertungen bis zur Spalte
  QuerSumme     As Long                 ' Quersumme der Wertungen bis zur Spalte
End Type



'------------------------------------------------------------------------------+
' Evaluate                                        - wende das Skatingsystem an |
' ---------------------------------------------------------------------------- |
' I/O: RundenWErtung()    - ein Feld mit allen Einzelwertungen                 |
' In : nJudges            - Anzahl der Wertungsrichter                         |
' ---------------------------------------------------------------------------- |
' Das Skatingsystem (die Auswertung der Platzierungen) wird angewendet. Dabei  |
' enthaelt das Feld RundenWertung fuer jeden Teilnehmer einen Eintrag des      |
' Typs TEinzelwertung, in dem die Startnummer und die Platzwertungen der       |
' Wertungsrichter eingetragen sind. Die Anzahl der Wertungsrichter wird ueber  |
' nJudges vorgegeben.                                                          |
' Nach der Anwendung des Skatingsystems werden die Platzierungen im Feld       |
' RundenWertung im jeweiligen Eintrag hinterlegt.                              |
'------------------------------------------------------------------------------+
Public Sub Evaluate(ByRef RundenWertung() As TEinzelwertung, _
                    ByVal nJudges As Byte)
  Dim nCountTeilnehmer  As Long                          ' Anzahl der Teilnehmer
  Dim actTeilnehmer     As Long                           ' aktueller Teilnehmer
  Dim nColumn           As Long                ' Spalte in der Auwertungs-Matrix
  Dim actJudge          As Long                      ' aktueller Wertungsrichter

  ' die anzahl der Teilnehmer ermitteln
  nCountTeilnehmer = UBound(RundenWertung, 1)

  ' die Auswertungs-Matrix aufspannen und alle Zellen initialisieren
  ReDim AuswertMatrix(nCountTeilnehmer, nCountTeilnehmer) As TMatrixZelle
  For actTeilnehmer = 1 To nCountTeilnehmer
    For nColumn = 1 To nCountTeilnehmer
      AuswertMatrix(actTeilnehmer, nColumn).Anzahl = 0
      AuswertMatrix(actTeilnehmer, nColumn).QuerSumme = 0
    Next nColumn
  Next actTeilnehmer

  ' das Feld der Platzierungen aufspannen und auf 0 setzen
  ReDim Platzierungen(nCountTeilnehmer) As Long
  For actTeilnehmer = 1 To nCountTeilnehmer
    Platzierungen(actTeilnehmer) = 0
  Next actTeilnehmer

  ' die Rundenwertung in die Auswertungs-Matrix uebertragen
  For actTeilnehmer = 1 To nCountTeilnehmer
    For actJudge = 1 To nJudges
      For nColumn = 1 To nCountTeilnehmer
        If (RundenWertung(actTeilnehmer).Wertung(actJudge) <= nColumn) Then
          AuswertMatrix(actTeilnehmer, nColumn).Anzahl = AuswertMatrix(actTeilnehmer, nColumn).Anzahl + 1
          AuswertMatrix(actTeilnehmer, nColumn).QuerSumme = AuswertMatrix(actTeilnehmer, nColumn).QuerSumme + RundenWertung(actTeilnehmer).Wertung(actJudge)
        End If
      Next nColumn
    Next actJudge
  Next actTeilnehmer

  ' jetzt das Skating-System anwenden
  Call EvalRule1((nJudges \ 2 + 1), 1, AuswertMatrix, Platzierungen)

  ' und die Platzierungen wieder zurueck in die Rundenwertung stecken
  For actTeilnehmer = 1 To nCountTeilnehmer
    RundenWertung(actTeilnehmer).Platz = Platzierungen(actTeilnehmer)
  Next actTeilnehmer

End Sub



'------------------------------------------------------------------------------+
' EvalRule1                                                 - wende Regel 1 an |
' ---------------------------------------------------------------------------- |
' In : aMajoritaet        - die geltende Majoritaet                            |
' I/O: aNextRank          - naechster zu vergebender Platz                     |
'      aAuswertMatrix()   - die aufbereitete Auswert-Matrix                    |
'      aPlatzierungen()   - Feld mit den Platzierungen                         |
' ---------------------------------------------------------------------------- |
' Ein Platz wird durch die absolute Mehrheit gleicher und ggf. besserer Wer-   |
' tungen entschieden.                                                          |
' Hat ein Teilnehmer die absolute Mehrheit fuer einen Platz, so erhaelt er     |
' diesen. Die gleichen Platzwertungen anderer Teilnehmer gelten dann fuer den  |
' naechstniederen Platz.                                                       |
' Haben zwei oder mehr Teilnehmer die absolute Mehrheit fuer einen Platz, so   |
' wird Regel 2 angewandt.                                                      |
'------------------------------------------------------------------------------+
Private Sub EvalRule1(ByVal aMajoritaet As Byte, _
                      ByRef aNextRank As Long, _
                      ByRef aAuswertMatrix() As TMatrixZelle, _
                      ByRef aPlatzierungen() As Long)
  Dim countTeilnehmer   As Long                          ' Anzahl der Teilnehmer
  Dim actTeilnehmer     As Long                        ' betrachteter Teilnehmer
  Dim actPosition       As Long                      ' betrachtete Matrix-Spalte
  Dim nDup              As Long      ' Anzahl des Auftretens gleicher Majoritaet
  Dim theOne            As Long     ' der Teilnehmer mit der eindeutigen Wertung

  countTeilnehmer = UBound(aAuswertMatrix, 1)  ' Anzahl der Teilnehmer ermitteln

  For actPosition = 1 To countTeilnehmer       ' alle Matrix-Spalten durchlaufen
    nDup = 0                               ' noch hat keiner hier die Majoritaet

    ' ermitteln, wieviele Teilnehmer in dieser Spalte die Majoritaet haben
    For actTeilnehmer = 1 To countTeilnehmer
      ' noch nicht bewertet und Majoritaet erreicht
      If ((aPlatzierungen(actTeilnehmer) = 0) And _
          (aAuswertMatrix(actTeilnehmer, actPosition).Anzahl >= aMajoritaet)) Then
        nDup = nDup + 1                              ' jetzt ein Teilnehmer mehr
        theOne = actTeilnehmer                                 ' naemlich dieser
      End If
    Next actTeilnehmer

    If (nDup > 1) Then ' mehrere Teilnehmer haben Majoritaet -> Regel 2 anwenden
      Call EvalRule2(countTeilnehmer, _
                     actPosition, _
                     aMajoritaet, _
                     aNextRank, _
                     aAuswertMatrix, _
                     aPlatzierungen)
    ElseIf (nDup = 1) Then                               ' Regel 1 ist eindeutig
      aPlatzierungen(theOne) = aNextRank                        ' Platz vergeben
      aNextRank = aNextRank + 1              ' und der naechste Platz kommt dran
    End If
  Next actPosition
End Sub



'------------------------------------------------------------------------------+
' EvalRule2                                                 - wende Regel 2 an |
' ---------------------------------------------------------------------------- |
' In : aCountTeilnehmer   - Anzahl der Teilnehmer                              |
'      aActPosition       - betrachtete Matrix-Spalte                          |
'      aMajoritaet        - die geltende Majoritaet                            |
' I/O: aNextRank          - naechster zu vergebender Platz                     |
'      aAuswertMatrix()   - die aufbereitete Auswert-Matrix                    |
'      aPlatzierungen()   - Feld mit den Platzierungen                         |
' ---------------------------------------------------------------------------- |
' Haben zwei oder mehr Teilnehmer die absolute Mehrheit fuer einen Platz, so   |
' erhaelt diesen der Teilnehmer mit der groesseren Mehrheit.                   |
' Haben zwei oder mehr Teilnehmer die gleiche absolute Mehrheit fuer einen     |
' Platz, so wird Regel 3 angewandt.                                            |
'------------------------------------------------------------------------------+
Private Sub EvalRule2(ByVal aCountTeilnehmer As Long, _
                      ByVal aActPosition As Long, _
                      ByVal aMajoritaet As Long, _
                      ByRef aNextRank As Long, _
                      ByRef aAuswertMatrix() As TMatrixZelle, _
                      ByRef aPlatzierungen() As Long)
  Dim nMaxMajoritaet  As Byte                      ' aktuell hoechste Majoritaet
  Dim nDup            As Long                            ' Anzahl des Auftretens
  Dim actTeilnehmer   As Long                          ' betrachteter Teilnehmer
  Dim theOne          As Long       ' der Teilnehmer mit der eindeutigen Wertung

  Do
    nMaxMajoritaet = 0                                        ' alles ist hoeher
    nDup = 0                                         ' aber noch nie aufgetreten

    ' zuerst hoechste Majoritaet und Anzahl Vorkommen ermitteln
    For actTeilnehmer = 1 To aCountTeilnehmer
      ' noch nicht bewertet und mindestens Majoritaet erreicht
      If ((aPlatzierungen(actTeilnehmer) = 0) And _
          (aAuswertMatrix(actTeilnehmer, aActPosition).Anzahl >= aMajoritaet)) Then
        ' hoechste Majoritaet ermitteln und Vorkommen zaehlen
        If (aAuswertMatrix(actTeilnehmer, aActPosition).Anzahl > nMaxMajoritaet) Then
          ' neue hoechste Majoritaet merken
          nMaxMajoritaet = aAuswertMatrix(actTeilnehmer, aActPosition).Anzahl
          nDup = 1                                            ' erstes Auftreten
          theOne = actTeilnehmer                         ' bei diesem Teilnehmer
        ElseIf (aAuswertMatrix(actTeilnehmer, aActPosition).Anzahl = nMaxMajoritaet) Then
          ' bereits bekannte Majoritaet
          nDup = nDup + 1                                   ' erneutes Auftreten
        End If
      End If
    Next actTeilnehmer

    ' und jetzt die Plaetze vergeben (oder tiefer graben)
    If (nDup > 1) Then      ' mehr als ein Teilnehmer gleich -> Regel 3 anwenden
      Call EvalRule3(aCountTeilnehmer, _
                     aActPosition, _
                     nMaxMajoritaet, _
                     aNextRank, _
                     aAuswertMatrix, _
                     aPlatzierungen)
    ElseIf (nDup = 1) Then                               ' Regel 2 ist eindeutig
      aPlatzierungen(theOne) = aNextRank                        ' Platz vergeben
      aNextRank = aNextRank + 1              ' und der naechste Platz kommt dran
    End If
  Loop Until (nDup = 0)           ' bis keine mehrfache hoechste Majoritaet mehr
End Sub



'------------------------------------------------------------------------------+
' EvalRule3                                                 - wende Regel 3 an |
' ---------------------------------------------------------------------------- |
' In : aCountTeilnehmer   - Anzahl der Teilnehmer                              |
'      aActPosition       - betrachtete Matrix-Spalte                          |
'      aMaxMajoritaet     - die geltende hoechste Majoritaet                   |
' I/O: aNextRank          - naechster zu vergebender Platz                     |
'      aAuswertMatrix()   - die aufbereitete Auswert-Matrix                    |
'      aPlatzierungen()   - Feld mit den Platzierungen                         |
' ---------------------------------------------------------------------------- |
' Haben zwei oder mehr Teilnehmer die gleiche absolute Mehrheit fuer einen     |
' Platz, so werden die Wertungen der Wertungsrichter, die diese Mehrheit bil-  |
' den, addiert. Der Teilnehmer mit der niedrigsten Summe erhaelt dann diesen   |
' Platz.                                                                       |
' Haben zwei oder mehr Teilnehmer die gleiche niedrigste Summe fuer einen      |
' Platz, so wird Regel 4 angewandt.                                            |
'------------------------------------------------------------------------------+
Private Sub EvalRule3(ByVal aCountTeilnehmer As Long, _
                      ByVal aActPosition As Long, _
                      ByVal aMaxMajoritaet As Byte, _
                      ByRef aNextRank As Long, _
                      ByRef aAuswertMatrix() As TMatrixZelle, _
                      ByRef aPlatzierungen() As Long)
  Const MAX_INT = 32768                            ' maximale Ausgangs-Quersumme

  Dim nMinQSumme      As Long                     ' aktuell niedrigste Quersumme
  Dim nDup            As Long                            ' Anzahl des Auftretens
  Dim actTeilnehmer   As Long               ' Laufvariable ueber alle Teilnehmer
  Dim theOne          As Long       ' der Teilnehmer mit der eindeutigen Wertung

  Do
    nMinQSumme = MAX_INT                                   ' alles ist niedriger
    nDup = 0                                         ' aber noch nie aufgetreten

    ' zuerst niedrigste Quersumme und Anzahl Vorkommen ermitteln
    For actTeilnehmer = 1 To aCountTeilnehmer
      ' noch nicht bewertet und hoechste Majoritaet (aMaxMajoritaet) erreicht
      If ((aPlatzierungen(actTeilnehmer) = 0) And _
          (aAuswertMatrix(actTeilnehmer, aActPosition).Anzahl = aMaxMajoritaet)) Then
        ' niedrigste Quersumme ermitteln und Vorkommen zaehlen
        If (aAuswertMatrix(actTeilnehmer, aActPosition).QuerSumme < nMinQSumme) Then
          ' neue niedrigste Quersumme merken
          nMinQSumme = aAuswertMatrix(actTeilnehmer, aActPosition).QuerSumme
          nDup = 1                                            ' erstes Auftreten
          theOne = actTeilnehmer                         ' bei diesem Teilnehmer
        ElseIf (aAuswertMatrix(actTeilnehmer, aActPosition).QuerSumme = nMinQSumme) Then
          ' bereits bekannte Quersumme
          nDup = nDup + 1                                   ' erneutes Auftreten
        End If
      End If
    Next actTeilnehmer
    ' und jetzt die Plaetze vergeben (oder tiefer graben)
    If (nDup > 1) Then      ' mehr als ein Teilnehmer gleich -> Regel 4 anwenden
      Call EvalRule4(aCountTeilnehmer, _
                     aActPosition, _
                     aMaxMajoritaet, _
                     nMinQSumme, _
                     aNextRank, _
                     aAuswertMatrix, _
                     aPlatzierungen)
    ElseIf (nDup = 1) Then                               ' Regel 3 ist eindeutig
      aPlatzierungen(theOne) = aNextRank                        ' Platz vergeben
      aNextRank = aNextRank + 1              ' und der naechste Platz kommt dran
    End If
 Loop Until (nMinQSumme = MAX_INT)        ' bis keine niedrigste Quersumme mehr

End Sub



'------------------------------------------------------------------------------+
' EvalRule4                                                 - wende Regel 4 an |
' ---------------------------------------------------------------------------- |
' In : aCountTeilnehmer   - Anzahl der Teilnehmer                              |
'      aActPosition       - betrachtete Matrix-Spalte                          |
'      aMaxMajoritaet     - die geltende hoechste Majoritaet                   |
'      aMinQSumme         - aktuell niedrigste Quersumme                       |
' I/O: aNextRank          - naechster zu vergebender Platz                     |
'      aAuswertMatrix()   - die aufbereitete Auswert-Matrix                    |
'      aPlatzierungen()   - Feld mit den Platzierungen                         |
' ---------------------------------------------------------------------------- |
' Haben zwei oder mehr Teilnehmer die gleiche niedrigste Summe fuer einen      |
' Platz, wird bei den betroffenen Teilnehmern der naechstniedrigere Platz      |
' (oder Plaetze, wenn notwendig) mit einbezogen.                               |
' Sind alle Wertungen beruecksichtigt und sind immer noch zwei oder mehr Teil- |
' nehmer gleich, so wird der Platz geteilt.                                    |
'------------------------------------------------------------------------------+
Private Sub EvalRule4(ByVal aCountTeilnehmer As Long, _
                      ByVal aActPosition As Long, _
                      ByVal aMaxMajoritaet As Byte, _
                      ByVal aMinQSumme As Long, _
                      ByRef aNextRank As Long, _
                      ByRef aAuswertMatrix() As TMatrixZelle, _
                      ByRef aPlatzierungen() As Long)
  Dim nSubPosition    As Long              ' aktuelle "niedrigere" Matrix-Spalte
  Dim nDup            As Long                            ' Anzahl des Auftretens
  Dim nLocalHigh      As Long                       ' lokales Maximum der Anzahl
  Dim actTeilnehmer   As Long               ' Laufvariable ueber alle Teilnehmer
  Dim nSplitRankCnt   As Long                 ' Anzahl gleich vergebener Plaetze
  Dim theOne          As Long       ' der Teilnehmer mit der eindeutigen Wertung
  Dim nLocalQuer      As Long       ' Quersumme des besten
  
    ' Start HK Dez. 2004
  Dim Platzgl(1 To 8) As Double ' Tabelle der platzgleichen Startnummern (maximal 7)
  Dim Platzg2(1 To 8) As Double ' Tabelle der platzgleichen Startnummern (maximal 7)
  nDup = 1
  Do While nDup < 8
    Platzgl(nDup) = 0
    Platzg2(nDup) = 0
    nDup = nDup + 1
    Loop
  nDup = 0
     ' Ende HK Dez. 2004
  If (aActPosition < aCountTeilnehmer) Then     ' schon ganz am Ende angekommen?

    ' nach rechts in Richtung niedrigere Plaetze laufen
    For nSubPosition = (aActPosition + 1) To aCountTeilnehmer
      ' und bei jedem niedrigeren Platz die Entscheidung suchen
      Do
        nLocalHigh = 0 ' alles ist hoeher
        nLocalQuer = 0 ' hk 07.06.04
        nDup = 0 ' aber noch nie aufgetreten

        For actTeilnehmer = 1 To aCountTeilnehmer
          If ((aPlatzierungen(actTeilnehmer) = 0) And _
              (aAuswertMatrix(actTeilnehmer, aActPosition).Anzahl = aMaxMajoritaet) And _
              (aAuswertMatrix(actTeilnehmer, aActPosition).QuerSumme = aMinQSumme)) Then
            If (aAuswertMatrix(actTeilnehmer, nSubPosition).Anzahl > nLocalHigh) Then
              ' neue hoechste Anzahl merken
              nLocalHigh = aAuswertMatrix(actTeilnehmer, nSubPosition).Anzahl
              ' If nLocalQuer = 0 Or nLocalQuer > aAuswertMatrix(actTeilnehmer, nSubPosition).QuerSumme Then  ' Neue Quersumme ist niedriger HK 07.06.04
              nLocalQuer = aAuswertMatrix(actTeilnehmer, nSubPosition).QuerSumme
              nDup = 1                                        ' erstes Auftreten
              theOne = actTeilnehmer                     ' bei diesem Teilnehmer
              ' End If
            ElseIf (aAuswertMatrix(actTeilnehmer, nSubPosition).Anzahl = nLocalHigh) Then ' 07.06.04 HK
              If nLocalQuer = aAuswertMatrix(actTeilnehmer, nSubPosition).QuerSumme Then   ' Quersumme ist gleich HK 07.06.04
              ' bereits bekannte hoechste Anzahl
                ' Start HK Dez. 2004
                If nDup = 1 Then
                    Platzgl(nDup) = theOne
                    nDup = nDup + 1                     ' ersten Teilnehmer in Tabelle
                End If
                Platzgl(nDup) = actTeilnehmer      ' zusätzliche Teilnehmer in Tabelle
                nDup = nDup + 1
                ' Ende HK Dez. 2004
              ElseIf nLocalQuer > aAuswertMatrix(actTeilnehmer, nSubPosition).QuerSumme Then   ' Quersumme ist kleiner HK 07.06.04
                ' Start HK Dez. 2004
                If nDup > 1 Then
                  Do While nDup > 0
                     Platzgl(nDup) = 0
                     nDup = nDup - 1
                     Loop
                 End If
                ' Ende HK Dez. 2004
                nLocalQuer = aAuswertMatrix(actTeilnehmer, nSubPosition).QuerSumme  ' 07.06.04 HK
                theOne = actTeilnehmer                      ' 07.06.04 HK
                nDup = 1                                    ' 07.06.04 HK
              End If
            End If
          End If
        Next actTeilnehmer
        If (nDup = 1) Then                                   ' Regel 4 eindeutig
          aPlatzierungen(theOne) = aNextRank                    ' Platz vergeben
          aNextRank = aNextRank + 1          ' und der naechste Platz kommt dran
          ' End                               ' HK 18.05.04
        End If
      Loop Until (nDup <> 1)                   ' bis Regel 4 nicht eindeutig ist
    Next nSubPosition
    If (nLocalHigh > 0) Then     ' es gibt am Ende immer noch gleiche Teilnehmer
      nSplitRankCnt = 0      ' noch keinmal geteilt
    'Start HK Dez. 2004
       Dim tabind As Double
       tabind = 0
       For tabind = 1 To nDup
            If Platzgl(tabind) = 0 Then
                tabind = nDup
            Else
              If ((aPlatzierungen(Platzgl(tabind)) = 0) And _
                (aAuswertMatrix(Platzgl(tabind), aActPosition).Anzahl = aMaxMajoritaet) And _
                (aAuswertMatrix(Platzgl(tabind), nSubPosition - 1).QuerSumme = nLocalQuer) And _
                (aAuswertMatrix(Platzgl(tabind), nSubPosition - 1).Anzahl = nLocalHigh)) Then ' HK 18.5.04
                aPlatzierungen(Platzgl(tabind)) = aNextRank             ' Platz vergeben
                nSplitRankCnt = nSplitRankCnt + 1                ' einmal mehr geteilt
               End If
            End If
        Next tabind
      ' Ende HK Dez. 2004
      aNextRank = aNextRank + nSplitRankCnt     ' geteilte Plaetze ueberspringen
    End If

  Else                           ' schon am Ende gestartet -> Platz wird geteilt

    nSplitRankCnt = 0                                     ' noch keinmal geteilt
    For actTeilnehmer = 1 To aCountTeilnehmer      ' alle Teilnehmer durchlaufen
      ' noch nicht bewertet und mit gleicher Majoritaet
      '     -----> hier muß noch getestet werden. Was ist, wenn am Ende immer noch
      '            Paare mit der gleichen Majorität und unterschiedlicher Additon
      '            der Plätze sind?
      If ((aPlatzierungen(actTeilnehmer) = 0) And _
          (aAuswertMatrix(actTeilnehmer, aActPosition).Anzahl = aMaxMajoritaet)) Then
        aPlatzierungen(actTeilnehmer) = aNextRank               ' Platz vergeben
        nSplitRankCnt = nSplitRankCnt + 1                  ' einmal mehr geteilt
      End If
    Next actTeilnehmer
    aNextRank = aNextRank + nSplitRankCnt       ' geteilte Plaetze ueberspringen

  End If
End Sub

Function msystem(RT_ID As Integer, tnr1, klasse, Runde, AnzahlWR, autoDeleteWertung As Boolean)
    
    Dim dbs As Database
    Set dbs = CurrentDb ' Bezug auf aktuelle Datenbank zurückgeben.
    Dim stmt As String
    
    Dim rstauswertung, rstSortiert, rstErstell, rstLoesch As Recordset
    Dim Anzahl, anzahln, Anzpaare As Double
    
    ' Anzahl Endrundenpaare ermitteln
    stmt = "select * from Paare_Rundenqualifikation where rt_id=" & RT_ID & " and anwesend_Status=1;"
    Set rstauswertung = dbs.OpenRecordset(stmt)
    If rstauswertung.EOF() Then
       MsgBox ("Keine Paare für diese Runde gefunden. Startkl:" & klasse & " Runde:" & Runde)
       End
    End If
    Set rstauswertung = rstauswertung.OpenRecordset()
    rstauswertung.MoveLast
    Anzpaare = rstauswertung.RecordCount
    stmt = "Select * from Auswertung a, Paare_Rundenqualifikation pr, Paare p, Wert_Richter wr where pr.rt_id=" & RT_ID & " and pr.tp_id=p.tp_id and a.wr_id=wr.wr_id and pr.pr_id=a.pr_id and a.platz>0 AND WR_AzuBi=False order by startnr, WR_Kuerzel"
    
    Set rstauswertung = dbs.OpenRecordset(stmt)
    If rstauswertung.EOF() Then
       MsgBox ("Keine Paare für diese Runde gefunden. Startkl:" & klasse & " Runde:" & Runde)
       End
    End If
    Set rstSortiert = rstauswertung.OpenRecordset()
    rstSortiert.MoveLast
    Anzahl = rstSortiert.RecordCount
    Dim thisRound() As modSkating.TEinzelwertung
    anzahln = Anzahl / AnzahlWR
    If anzahln <> Anzpaare Then
       MsgBox ("Noch nicht alle Wertungen eingegeben. Bisher erst " & anzahln)
       End
    End If
    ReDim Preserve thisRound(anzahln)
    rstSortiert.MoveFirst
    Dim snr, ind1, ind2
    ind1 = 1
    ind2 = 1
    snr = 0
    snr = rstSortiert!Startnr
    thisRound(ind1).StartNo = rstSortiert!Startnr
    thisRound(ind1).TP_ID = rstSortiert![p.TP_ID]
    
    Do Until rstSortiert.EOF
       If snr = rstSortiert!Startnr Then
          thisRound(ind1).Wertung(ind2) = rstSortiert![a.Platz]
          thisRound(ind1).Punkt(ind2) = rstSortiert![a.Punkte]
          thisRound(ind1).WR_ID(ind2) = rstSortiert![a.wr_id]
          ind2 = ind2 + 1
          rstSortiert.MoveNext
        Else
          ind2 = 1
          ind1 = ind1 + 1
          snr = rstSortiert!Startnr
          thisRound(ind1).StartNo = rstSortiert!Startnr
          thisRound(ind1).TP_ID = rstSortiert![p.TP_ID]
    ' Sortieren (09.09.2003) HK
        End If
       Loop
    anzahln = ind1
    Call modSkating.Evaluate(thisRound, AnzahlWR)
    Set rstLoesch = dbs.OpenRecordset("select * from Majoritaet where rt_id=" & RT_ID)
    If Not rstLoesch.EOF() Then
     With rstLoesch
         rstLoesch.MoveLast
         If rstLoesch.RecordCount > 0 Then
            Dim Mldg, Stil, titel, Antwort, Text1
            Mldg = "Majorität wurde schon errechnet. Bestehende Auswertung löschen?"   ' Meldung definieren.
            Stil = vbYesNo + vbCritical + vbDefaultButton2  ' Schaltflächen definieren.
            titel = "Was nun?"  ' Titel definieren.
            If (autoDeleteWertung) Then
                Antwort = vbYes
            Else
                Antwort = MsgBox(Mldg, Stil, titel)    ' Meldung anzeigen.
            End If
            If (Antwort = vbYes) Then ' Benutzer hat "Ja" gewählt.
               rstLoesch.MoveFirst
               Do Until rstLoesch.EOF
                  .Delete
                  rstLoesch.MoveNext
                  Loop
             Else
                rstauswertung.Close
                rstSortiert.Close
                End
             End If
          End If
    End With
    End If
    
    Set rstErstell = dbs.OpenRecordset("Majoritaet", dbOpenDynaset)
    With rstErstell
    For snr = 1 To anzahln
        .AddNew
        !RT_ID = RT_ID
        !TP_ID = thisRound(snr).TP_ID
        !WR1 = thisRound(snr).Wertung(1)
        !WR2 = thisRound(snr).Wertung(2)
        !WR3 = thisRound(snr).Wertung(3)
        !WR4 = thisRound(snr).Wertung(4)
        !WR5 = thisRound(snr).Wertung(5)
        !WR6 = thisRound(snr).Wertung(6)
        !WR7 = thisRound(snr).Wertung(7)
        !WR1_Orig_Platz = thisRound(snr).Wertung(1)
        !WR2_Orig_Platz = thisRound(snr).Wertung(2)
        !WR3_Orig_Platz = thisRound(snr).Wertung(3)
        !WR4_Orig_Platz = thisRound(snr).Wertung(4)
        !WR5_Orig_Platz = thisRound(snr).Wertung(5)
        !WR6_Orig_Platz = thisRound(snr).Wertung(6)
        !WR7_Orig_Platz = thisRound(snr).Wertung(7)
        !WR1_Orig_Punkte = thisRound(snr).Punkt(1)
        !WR2_Orig_Punkte = thisRound(snr).Punkt(2)
        !WR3_Orig_Punkte = thisRound(snr).Punkt(3)
        !WR4_Orig_Punkte = thisRound(snr).Punkt(4)
        !WR5_Orig_Punkte = thisRound(snr).Punkt(5)
        !WR6_Orig_Punkte = thisRound(snr).Punkt(6)
        !WR7_Orig_Punkte = thisRound(snr).Punkt(7)
        !wr1_Platz = thisRound(snr).Wertung(1)
        !wr2_Platz = thisRound(snr).Wertung(2)
        !wr3_Platz = thisRound(snr).Wertung(3)
        !wr4_Platz = thisRound(snr).Wertung(4)
        !wr5_Platz = thisRound(snr).Wertung(5)
        !wr6_Platz = thisRound(snr).Wertung(6)
        !wr7_Platz = thisRound(snr).Wertung(7)
        !WR1_Punkte = thisRound(snr).Punkt(1)
        !WR2_Punkte = thisRound(snr).Punkt(2)
        !WR3_Punkte = thisRound(snr).Punkt(3)
        !wr4_Punkte = thisRound(snr).Punkt(4)
        !WR5_Punkte = thisRound(snr).Punkt(5)
        !WR6_Punkte = thisRound(snr).Punkt(6)
        !WR7_Punkte = thisRound(snr).Punkt(7)
        !wr1_id = thisRound(snr).WR_ID(1)
        !wr2_id = thisRound(snr).WR_ID(2)
        !wr3_id = thisRound(snr).WR_ID(3)
        !wr4_id = thisRound(snr).WR_ID(4)
        !wr5_id = thisRound(snr).WR_ID(5)
        !wr6_id = thisRound(snr).WR_ID(6)
        !wr7_id = thisRound(snr).WR_ID(7)
        !Platz = thisRound(snr).Platz
        !Platz_Orig = thisRound(snr).Platz
        .Update
        .Bookmark = .LastModified
      Next snr
    End With
    rstauswertung.Close
    rstSortiert.Close
    rstErstell.Close
    Set dbs = Nothing
End Function

Sub PaareInDieNaechsteRunde(Turniernr As Integer, currentRT_ID As Integer, nextRT_ID As Integer, bisPlatz As Integer, Rundentext As String)
    Call PaareInDieNaechsteRunde2(Turniernr, currentRT_ID, nextRT_ID, 1, bisPlatz, Rundentext)
End Sub

Sub PaareInDieNaechsteRunde2(Turniernr As Integer, currentRT_ID As Integer, nextRT_ID As Integer, vonPlatz As Integer, bisPlatz As Integer, Rundentext As String)
    Dim dbs As Database
    Dim rstmajoritaet, rstquali, rstpaare As Recordset
    Dim re As Recordset
    Dim stmt As String
    
    Set dbs = CurrentDb()
    Set rstpaare = dbs.OpenRecordset("Paare", dbOpenDynaset)
    Set rstquali = dbs.OpenRecordset("Paare_rundenqualifikation", dbOpenDynaset)
    Dim naechste_Rd As String
    naechste_Rd = "" ' TODO Forms.Majoritaet_ausrechnen!nächste_Runde
    
    stmt = "select * from Majoritaet where RT_ID=" & currentRT_ID & " and platz <= " & bisPlatz & " and platz>=" & vonPlatz
    
    Set rstmajoritaet = dbs.OpenRecordset(stmt)
    If Not rstmajoritaet.EOF() Then
        ' HK 13.11.2011 Paare in der Runde löschen, wenn die Runde schon besteht.
                rstquali.FindFirst ("rt_id=" & nextRT_ID)
                If Not rstquali.NoMatch Then
                    Dim result As Integer
                    result = MsgBox("Die Runde besteht schon, sollen die Paare in dieser Runde gelöscht werden?", vbYesNo)
                    If (result = vbYes) Then
                        dbs.Execute ("Delete from Paare_rundenqualifikation where rt_id=" & nextRT_ID)
                    End If
                End If
        '                     '
        Do While Not rstmajoritaet.EOF()
            rstpaare.FindFirst ("TP_ID=" & rstmajoritaet!TP_ID)
            If Not rstpaare.NoMatch Then
                rstquali.FindFirst ("rt_id=" & nextRT_ID & " and tp_id=" & rstmajoritaet!TP_ID)
                
                If rstquali.NoMatch Then
                    rstquali.AddNew
                Else
                    rstquali.Edit
                End If
                
                rstquali!RT_ID = nextRT_ID
                rstquali!TP_ID = rstmajoritaet!TP_ID
                rstquali!Rundennummer = Null
                rstquali!Anwesend_Status = 1 'rstpaare!Anwesent_Status
                ' rstquali!Verein_Name = rstpaare!Verein_Name  ' für Access97 aktivieren
                rstquali!Verein_Name = Replace(rstpaare!Verein_Name, "'", "§")      ' Für access XP aktivieren
                rstquali!Verein_Name = Replace(rstquali!Verein_Name, "`", "§")      ' Für access XP aktivieren
                rstquali!Verein_Name = Replace(rstquali!Verein_Name, "´", "§")      ' Für access XP aktivieren
                rstquali!Verein_Name = Replace(rstquali!Verein_Name, Chr(34), "§")  ' Für access XP aktivieren " in § umwandeln
                rstquali.Update
                
                rstpaare.Edit
                rstpaare!RT_ID_Ausgeschieden = Null
                rstpaare!Platz = 0
                rstpaare!Punkte = 0
                rstpaare.Update
                
                rstmajoritaet.Edit
                rstmajoritaet!rt_id_weiter = nextRT_ID
                rstmajoritaet.Update
                
            Else
                MsgBox ("Paar in der Tabelle Paare nicht gefunden. Startnummer: " & rstmajoritaet!Startnr)
                End
            End If
            rstmajoritaet.MoveNext
        Loop
        make_a_startlist nextRT_ID
        ' bei zweigeteilten Runden 2. Runde einfügen
        Set re = dbs.OpenRecordset("SELECT * FROM Rundentab WHERE RT_ID=" & nextRT_ID & ";")
            Dim rs As Recordset
            Dim rtida As Integer
            Dim rtidf As Integer
            Dim sqlstr As String
            Dim sk As String
            'RnR
            If (re!Startklasse = "RR_A" Or re!Startklasse = "RR_B") And InStr(1, re!Runde, "End_r_") > 0 Then
                sqlstr = "select * from Rundentab where startklasse = '" & re!Startklasse & "' and turniernr = " & Turniernr & " and runde='End_r_Fuß'"
                Set rs = dbs.OpenRecordset(sqlstr)
                ' wird ziel, quelle, neuer rt_id
                dbs.Execute ("Delete from Paare_rundenqualifikation where rt_id=" & rs!RT_ID)
                Set rstmajoritaet = dbs.OpenRecordset("SELECT * FROM Paare_Rundenqualifikation WHERE RT_ID=" & re!RT_ID & ";")
                fill_Paare_rundenquali rstquali, rstmajoritaet, rs!RT_ID
                make_a_startlist rs!RT_ID
                rs.Close
            End If
            'Boogie
            If (re!Startklasse = "BW_SA" Or re!Startklasse = "BW_MA") And InStr(1, re!Runde, "_r_schnell") > 0 Then
                sqlstr = "select * from Rundentab where startklasse = '" & re!Startklasse & "' and turniernr = " & Turniernr & " and runde='" & left(re!Runde, 3) & "_r_lang'"
                Set rs = dbs.OpenRecordset(sqlstr)
                ' wird quelle
                dbs.Execute ("Delete from Paare_rundenqualifikation where rt_id=" & rs!RT_ID)
                Set rstmajoritaet = dbs.OpenRecordset("SELECT * FROM Paare_Rundenqualifikation WHERE RT_ID=" & re!RT_ID & ";")
                fill_Paare_rundenquali rstquali, rstmajoritaet, rs!RT_ID
                make_a_startlist rs!RT_ID
                rs.Close
            End If
            'Breitensport Baden-Württemberg
            If get_bs_erg(re!Startklasse, 3) = "BWBS_" And InStr(1, re!Runde, "End_r_2") Then
                sqlstr = "select * from Rundentab where startklasse = '" & re!Startklasse & "' and turniernr = " & Turniernr & " and runde='End_r_1'"
                Set rs = dbs.OpenRecordset(sqlstr)
                ' wird quelle
                dbs.Execute ("Delete from Paare_rundenqualifikation where rt_id=" & rs!RT_ID)
                Set rstmajoritaet = dbs.OpenRecordset("SELECT * FROM Paare_Rundenqualifikation WHERE RT_ID=" & re!RT_ID & ";")
                fill_Paare_rundenquali rstquali, rstmajoritaet, rs!RT_ID
                make_a_startlist rs!RT_ID
                rs.Close
                
            End If
    End If ' HK 07.11.11
    Call WriteRundeReport(currentRT_ID)

End Sub

' ------------------------------------------------------------------------
' Füllt das Attribut Runde_Report in der Tabelle Majoritaet für die Runde
' mit der Id RT_ID. Dies ist notwendig, da ansonsten die Trennlinien
' bei der Reports für die Ergebnislisten der einzelnen Runden nicht richtig
' darstellbar sind.
' ------------------------------------------------------------------------
Public Sub WriteRundeReport(RT_ID As Integer)
    Dim dbs As Database
    Set dbs = CurrentDb

    Dim rst As Recordset
    Set rst = dbs.OpenRecordset("Select * from Majoritaet where RT_ID=" & RT_ID & " order by Platz")
    If Not rst.EOF Then rst.MoveFirst
    
    Dim lastRT_ID As Integer
    Dim currentRT_ID As Integer
    Dim rundeReport As Integer
    lastRT_ID = -1
    rundeReport = -1
    
    Do While Not rst.EOF
        If (IsNull(rst!rt_id_weiter)) Then
            currentRT_ID = 0
        Else
            currentRT_ID = rst!rt_id_weiter
        End If
        
        If (currentRT_ID <> lastRT_ID) Then
            lastRT_ID = currentRT_ID
            rundeReport = rundeReport + 1
        End If
        
        rst.Edit
        rst!Runde_Report = rundeReport
        rst.Update
        
        rst.MoveNext
    Loop
    
    rst.Close
End Sub

' ------------------------------------------------------------------------
' Füllt das Attribut Runde_Report in der Tabelle Paare.
' Dies ist notwendig, da ansonsten die Trennlinien bei der Reports für die
' Ergebnislisten der einzelnen Runden nicht richtig darstellbar sind.
' ------------------------------------------------------------------------
Public Sub WriteRundeReport_Paare()
    Dim dbs As Database
    Set dbs = CurrentDb

    Dim rst As Recordset
    Set rst = dbs.OpenRecordset("Select * from Paare where platz<>0 order by turniernr, startkl, platz")
    rst.MoveFirst
    
    Dim lastRT_ID As Integer
    Dim currentRT_ID As Integer
    Dim rundeReport As Integer
    lastRT_ID = -1
    rundeReport = -1
    
    Do While Not rst.EOF
        If (IsNull(rst!RT_ID_Ausgeschieden)) Then
            currentRT_ID = 0
        Else
            currentRT_ID = rst!RT_ID_Ausgeschieden
        End If
        
        If (currentRT_ID <> lastRT_ID) Then
            lastRT_ID = currentRT_ID
            rundeReport = rundeReport + 1
        End If
        
        rst.Edit
        rst!Runde_Report = rundeReport
        rst.Update
        
        rst.MoveNext
    Loop
    
    rst.Close
End Sub

Public Function GetPaareInRunde(RT_ID As Integer) As Integer
    Dim dbs As Database
    Set dbs = CurrentDb

    Dim rst As Recordset
    Set rst = dbs.OpenRecordset("Select count(*) as Anzahl from Paare_Rundenqualifikation pr where rt_id=" & RT_ID)
    GetPaareInRunde = rst!Anzahl
    rst.Close
End Function

Public Function GetPaareBisPlatz(RT_ID As Integer, Platz As Integer) As Integer
    Dim dbs As Database
    Set dbs = CurrentDb

    Dim rst As Recordset
    Dim stmt As String
    stmt = "Select count(*) as Anzahl from Majoritaet m where rt_id=" & RT_ID & " and platz<=" & Platz
    Set rst = dbs.OpenRecordset(stmt)
    GetPaareBisPlatz = rst!Anzahl
    rst.Close
End Function

Public Sub PaarePlatzieren(RT_ID As Integer, vonPlatz As Integer)
    Call PaarePlatzierenMitHoffnungsrunde(RT_ID, vonPlatz, 0)
End Sub

Public Sub PaarePlatzierenMitHoffnungsrunde(RT_ID As Integer, vonPlatz As Integer, offsetPlatz As Integer)
    Dim dbs As Database
    Set dbs = CurrentDb
    Dim Punkte As Integer
    
    Dim rstmajoritaet1, rstpaare1, HatAufstiegspunkte As Recordset
    Set rstmajoritaet1 = dbs.OpenRecordset("select * from majoritaet where rt_id=" & RT_ID & " and Platz >= " & vonPlatz & " order by platz desc")
    If (Not rstmajoritaet1.EOF()) Then
    
        rstmajoritaet1.MoveFirst
        Set rstpaare1 = dbs.OpenRecordset("Paare")
        
        '*****AB***** V14.10 für die Tanzrunde die Startklasse ermitteln, ob diese Aufstiegspunkte hat
        Set HatAufstiegspunkte = dbs.OpenRecordset("SELECT Rundentab.RT_ID, Startklasse.hatAufstiegspunkte FROM Startklasse INNER JOIN Rundentab ON Startklasse.Startklasse = Rundentab.Startklasse WHERE (((Rundentab.RT_ID)=" & RT_ID & "));")
        HatAufstiegspunkte.MoveFirst
        
        Do While Not rstmajoritaet1.EOF()
            rstpaare1.FindFirst ("TP_ID=" & rstmajoritaet1!TP_ID)
            
            If Not rstpaare1.NoMatch Then
                 rstpaare1.Edit
                 
                 rstpaare1!Platz = rstmajoritaet1!Platz + offsetPlatz
                
                '*****AB***** V14.10 Startklassen ohne Aufstiegspunkte auf NULL setzen
                 If HatAufstiegspunkte!HatAufstiegspunkte Then
                    rstpaare1!Punkte = getPunkte(rstpaare1!Turniernr, rstpaare1!Startkl, rstpaare1!TP_ID, RT_ID, rstpaare1!Platz, rstmajoritaet1!DQ_ID, offsetPlatz, False) + IIf(left(rstpaare1!Startkl, 3) = "RR_", 10, 0)
                 Else
                    rstpaare1!Punkte = Null
                 End If
                 rstpaare1!Ranglistenpunkte = getPunkte(rstpaare1!Turniernr, rstpaare1!Startkl, rstpaare1!TP_ID, RT_ID, rstpaare1!Platz, rstmajoritaet1!DQ_ID, offsetPlatz, False)
                                                                          
                 rstpaare1!RT_ID_Ausgeschieden = RT_ID
                 
                 rstpaare1.Update
            Else
                MsgBox ("Paar nicht gefunden: " & rstmajoritaet1!Startnr)
            End If
            rstmajoritaet1.MoveNext
        Loop
        rstpaare1.Close
        rstmajoritaet1.Close
        Call WriteRundeReport_Paare
        
    End If

End Sub

Public Function getPunkte(Turniernr As Integer, Startklasse As String, TP_ID As Integer, RT_ID As Integer, Platz As Integer, DQ_ID As Integer, offsetPlatz As Integer, isAufstiegspunkte As Boolean) As Double

    ' Anzahl der Paare ermitteln, die Platzgleich sind
    Dim dbs As Database
    Dim rst As Recordset
    Dim rstmajoritaet As Recordset
    Dim stmt As String
    Dim anzahlPaareAufPlatz As Integer
    
    Set dbs = CurrentDb
    
    stmt = "select count(*) as anzahl from majoritaet where rt_id=" & RT_ID & " and Platz=" & (Platz - offsetPlatz)
    Set rstmajoritaet = dbs.OpenRecordset(stmt)
    anzahlPaareAufPlatz = rstmajoritaet!Anzahl
    rstmajoritaet.Close
    
    getPunkte = getPunkteFuerPlatz(Platz, anzahlPaareAufPlatz, isAufstiegspunkte)
    
    ' Wenn eine Disqualifikation in der ersten Runde erfolgt ist, dann
    ' gibts 0 Punkte
    stmt = "select count(*) as Anzahl from Paare where turniernr=" & Turniernr & " and Anwesent_Status=1 and Startkl='" & Startklasse & "' and platz>" & (Platz - 1)
    Set rst = dbs.OpenRecordset(stmt)
    getPunkte = getPunkte + IIf(isAufstiegspunkte, 0, rst!Anzahl)
    
    If (rst!Anzahl = 0 And DQ_ID > 0) Then
        getPunkte = 0
    End If
    rst.Close
    ' Wenn Punkte für den Aufstieg, dann nur in den entsprechenden Startklassen
    If (isAufstiegspunkte) Then
        stmt = "Select * from Startklasse where Startklasse='" & Startklasse & "'"
        Set rst = dbs.OpenRecordset(stmt)
        Dim HatAufstiegspunkte As Integer
        HatAufstiegspunkte = rst!HatAufstiegspunkte
        rst.Close
        
        If (Not HatAufstiegspunkte) Then
            getPunkte = 0
        End If
    End If
End Function

Public Function getPunkteFuerPlatz(Platz As Integer, anzahlPaareAufPlatz As Integer, isAufstiegspunkte As Boolean) As Double
    Dim anz As Integer
    Dim GesamtPunkte As Integer
    Dim Punkte As Integer
    anz = anzahlPaareAufPlatz
    GesamtPunkte = 0
    Do While (anz > 0)
        If isAufstiegspunkte Then
            Punkte = getAufstiegPuPlatz(Platz)
        Else
            Punkte = getPunkteVonPlatz(Platz)
        End If
        GesamtPunkte = GesamtPunkte + Punkte
        Platz = Platz + 1
        anz = anz - 1
    Loop
    
    getPunkteFuerPlatz = GesamtPunkte / anzahlPaareAufPlatz
    
    ' Aufstiegspunkte auf ganze Punkte und Ranglistenpunkte auf zwei Stellen nach dem Komma runden
    If (isAufstiegspunkte) Then
        getPunkteFuerPlatz = Round(getPunkteFuerPlatz)
    Else
        getPunkteFuerPlatz = Round(getPunkteFuerPlatz, 2)
    End If

End Function

Private Function getPunkteVonPlatz(Platz As Integer) As Double
    If (Platz = 1) Then
        getPunkteVonPlatz = 20
    ElseIf (Platz = 2) Then
        getPunkteVonPlatz = 15
    ElseIf (Platz = 3) Then
        getPunkteVonPlatz = 10
    ElseIf (Platz = 4) Then
        getPunkteVonPlatz = 8
    ElseIf (Platz = 5) Then
        getPunkteVonPlatz = 6
    ElseIf (Platz = 6) Then
        getPunkteVonPlatz = 4
    ElseIf (Platz = 7) Then
        getPunkteVonPlatz = 2
    ElseIf (Platz = 8) Then
        getPunkteVonPlatz = 1
    Else
        getPunkteVonPlatz = 0
    End If
End Function

Private Function getAufstiegPuPlatz(Platz As Integer) As Double
    If (Platz = 1) Then
        getAufstiegPuPlatz = 120
    ElseIf (Platz = 2) Then
        getAufstiegPuPlatz = 100
    ElseIf (Platz = 3) Then
        getAufstiegPuPlatz = 90
    ElseIf (Platz <= 20) Then
        getAufstiegPuPlatz = 84 - (Platz - 4) * 4
    ElseIf (Platz <= 29) Then
        getAufstiegPuPlatz = 18 - (Platz - 21) * 2
    Else
        getAufstiegPuPlatz = 0
    End If
End Function

Public Sub UpdateAnzahl_Paare(RT_ID As Integer)
    Dim Anzahl As Integer
    Dim dbs As Database
    Set dbs = CurrentDb
    
    Dim rst As Recordset
    Set rst = dbs.OpenRecordset("select count(*) as Anzahl from Majoritaet where rt_id=" & RT_ID)
    Anzahl = rst!Anzahl
    rst.Close
    
    Set rst = dbs.OpenRecordset("Select * from rundentab where rt_id=" & RT_ID)
    Dim rstAnzahlPaare As Recordset
    Dim Runde As String
    Runde = rst!Runde
    Set rstAnzahlPaare = dbs.OpenRecordset("Anzahl_Paare")
    Dim search As String
    search = "Turniernr=" & rst!Turniernr & " and Klasse='" & rst!Startklasse & "'"
    rstAnzahlPaare.FindFirst (search)
    
    If (rstAnzahlPaare.NoMatch) Then
        rstAnzahlPaare.AddNew
        rstAnzahlPaare!klasse = rst!Startklasse
        rstAnzahlPaare!Turniernr = rst!Turniernr
    Else
        rstAnzahlPaare.Edit
    End If
    
    If (Runde = "Vor_r" Or Runde = "Vor_r_Akro" Or Runde = "Vor_r_schnell") Then
        rstAnzahlPaare!VR = Anzahl
    ElseIf (Runde = "Hoff_r") Then
        rstAnzahlPaare!hr = Anzahl
    ElseIf (Runde = "1_Zw_r") Then
        rstAnzahlPaare!erstzr = Anzahl
    ElseIf (Runde = "2_Zw_r") Then
        rstAnzahlPaare!zweitzr = Anzahl
    ElseIf (Runde = "Stich_r") Then
        rstAnzahlPaare!stichr = Anzahl
    ElseIf (Runde = "Stich_r_1pl") Then
        rstAnzahlPaare!stichr1pl = Anzahl
    ElseIf (Runde = "End_r" Or Runde = "End_r_Akro" Or Runde = "End_r_2" Or Runde = "End_r_schnell") Then
        rstAnzahlPaare!er = Anzahl
    End If
    
    rstAnzahlPaare.Update
End Sub

Public Sub PaareInRundeNachPunktabzugPlatzieren(RT_ID As Integer, TP_ID As Integer, Turniernr As Integer, Startkl As String, Runde As String, AnzahlWR As Integer, Anzahl_Abzuege As Integer)

    Dim dbs As Database
    Dim rst1 As Recordset
    Dim Punktabzug As Double
    Dim stmt1 As String
    
    Set dbs = CurrentDb
    Punktabzug = get_verstoss(Startkl)
    
    'Zunächst alle Punkte auf den richtigen Endwert setzen
    
    Dim Platz As Integer
    Dim P, Abzuege As Double
    stmt1 = "Select * from majoritaet where rt_id=" & RT_ID & " and tp_id=" & TP_ID
    Set rst1 = dbs.OpenRecordset(stmt1)
    Dim i As Integer
    Dim wrp_orig, wrp As String
    ' Das Paar wird sofort aufgerufen, somit "Do While" mit "Loop" entfallen
    ' Do While (Not rst.EOF)
        rst1.Edit
        ' Die Punkteabzüge einarbeiten
        Abzuege = getAnzahlPunktanzug(Anzahl_Abzuege) * Punktabzug
        
        rst1![WR1_Punkte] = getReducedPunkte(rst1!WR1_Orig_Punkte, Abzuege)
        rst1![WR2_Punkte] = getReducedPunkte(rst1!WR2_Orig_Punkte, Abzuege)
        rst1![WR3_Punkte] = getReducedPunkte(rst1!WR3_Orig_Punkte, Abzuege)
        rst1![wr4_Punkte] = getReducedPunkte(rst1!WR4_Orig_Punkte, Abzuege)
        rst1![WR5_Punkte] = getReducedPunkte(rst1!WR5_Orig_Punkte, Abzuege)
        rst1![WR6_Punkte] = getReducedPunkte(rst1!WR6_Orig_Punkte, Abzuege)
        rst1![WR7_Punkte] = getReducedPunkte(rst1!WR7_Orig_Punkte, Abzuege)
        rst1!PA_ID = Anzahl_Abzuege   'HK 20111201
        rst1.Update
    '    rst1.MoveNext
    ' Loop
    
    rst1.Close
    
    ' Jetzt die Plätze bei den Wertungsrichtern anpassen
    Dim letztePunktzahl As Double
    Dim anzahlPaareAufPlatz As Integer
    Dim letzterOrigPlatz As Integer
    
    Dim wrnr As Integer
    For wrnr = 1 To AnzahlWR
        
        Platz = 0
        letztePunktzahl = -1
        anzahlPaareAufPlatz = 1
        letzterOrigPlatz = -1
        
        stmt1 = "Select * from majoritaet where rt_id=" & RT_ID & " order by WR" & wrnr & "_Punkte desc, WR" & wrnr & "_Orig_Platz asc"
        Set rst1 = dbs.OpenRecordset(stmt1)
           
        Do While (Not rst1.EOF)
            rst1.Edit
            
            If ((letztePunktzahl <> rst1("WR" & wrnr & "_Punkte")) Or (letzterOrigPlatz <> rst1("WR" & wrnr & "_Orig_Platz"))) Then
                Platz = Platz + anzahlPaareAufPlatz
                anzahlPaareAufPlatz = 1
                letztePunktzahl = rst1("WR" & wrnr & "_Punkte")
                letzterOrigPlatz = rst1("WR" & wrnr & "_Orig_Platz")
            Else
                anzahlPaareAufPlatz = anzahlPaareAufPlatz + 1
            End If
            
            rst1("WR" & wrnr & "_Platz") = Platz
            rst1("WR" & wrnr) = Platz
            
            rst1.Update
            rst1.MoveNext
        Loop
        rst1.Close
    
    Next
    
    ' Jetzt über alle möglichen WR iterieren und die Platzierungen korigieren
    Dim wrNum As Integer
    For wrNum = 1 To AnzahlWR
        Call wr_plaetze_korrigieren(RT_ID, "WR" & wrNum)
    Next
    
    ' Anzahl der Tanzpaare in dieser Runde ermitteln
    Dim anzahlPaare As Integer
    
    stmt1 = "Select count(*) as Anzahl from majoritaet where rt_id=" & RT_ID
    Set rst1 = dbs.OpenRecordset(stmt1)
    rst1.MoveFirst
    anzahlPaare = rst1!Anzahl
    rst1.Close
    
    ' Die neuen Platzierungen der Paare in ein Array Schreiben
    Dim thisRound() As TEinzelwertung
    ReDim Preserve thisRound(anzahlPaare)
    Dim pos As Integer
    pos = 1
    
    stmt1 = "Select * from majoritaet where rt_id=" & RT_ID
    Set rst1 = dbs.OpenRecordset(stmt1)
    Do While Not rst1.EOF
        
        thisRound(pos).TP_ID = rst1!TP_ID
        thisRound(pos).Wertung(1) = rst1!WR1
        thisRound(pos).Wertung(2) = rst1!WR2
        thisRound(pos).Wertung(3) = rst1!WR3
        thisRound(pos).Wertung(4) = rst1!WR4
        thisRound(pos).Wertung(5) = rst1!WR5
        thisRound(pos).Wertung(6) = rst1!WR6
        thisRound(pos).Wertung(7) = rst1!WR7
        
        pos = pos + 1
        rst1.MoveNext
    Loop
    
    rst1.Close
    
    ' Majoritätswertung durchführen
    Call Evaluate(thisRound, AnzahlWR)
    
    ' Ergebnis der Majoritätswertung wieder auf die Tanzpaare übertragen
    stmt1 = "Select * from Majoritaet where RT_ID=" & RT_ID
    Set rst1 = dbs.OpenRecordset(stmt1, dbOpenDynaset)
    For pos = 1 To anzahlPaare
        rst1.FindFirst ("TP_ID=" & thisRound(pos).TP_ID)
        rst1.Edit
        rst1!Platz = thisRound(pos).Platz
        rst1.Update
    Next
    rst1.Close
    
End Sub

Private Function getReducedPunkte(Punkte_Orig As Double, Abzuege As Double)
    Dim result As Double
    result = Punkte_Orig - Abzuege
    If (result < 0) Then
        result = 0
    End If
    getReducedPunkte = result
End Function

Private Function getAnzahlPunktanzug(PA_ID As Integer)
    Dim dbs As Database
    Set dbs = CurrentDb
    
    Dim rst As Recordset
    Dim stmt As String
    
    stmt = "Select * from Punktabzug where pa_id=" & PA_ID
    
    Set rst = dbs.OpenRecordset(stmt)
    rst.MoveFirst
    getAnzahlPunktanzug = rst!PA_ANZAHL
    rst.Close
    
End Function

Private Sub wr_plaetze_korrigieren(RT_ID As Integer, wertungsrichter As String)
    
    Dim dbs As Database
    Set dbs = CurrentDb
    
    Dim rst As Recordset
    Dim stmt As String
    
    ' Jetzt alle Plätze um die disqualifizierten Paare korrigieren
    Dim anzDQ As Integer            ' Anzahl der Disqualifikationen
    Dim anzDQaktPlatz As Integer    ' Anzahl der Disqualifikationen auf dem aktuellen Platz
    Dim aktPlatz As Integer         ' Aktueller Platz
    Dim anzPaareOhneDQ As Integer   ' Anzahl der Paare ohne Disqualifikation
    anzDQ = 0
    anzDQaktPlatz = 0
    aktPlatz = 0
    anzPaareOhneDQ = 0
    
    stmt = "Select * from majoritaet where rt_id=" & RT_ID & " order by " & wertungsrichter
    Set rst = dbs.OpenRecordset(stmt)
    Do While (Not rst.EOF)
        If (rst(wertungsrichter) <> aktPlatz) Then
            anzDQ = anzDQ + anzDQaktPlatz
            aktPlatz = rst(wertungsrichter)
            anzDQaktPlatz = 0
        End If
        
        If (rst!DQ_ID > 0) Then
            anzDQaktPlatz = anzDQaktPlatz + 1
        Else
            rst.Edit
            rst(wertungsrichter) = rst(wertungsrichter) - anzDQ
            rst.Update
            anzPaareOhneDQ = anzPaareOhneDQ + 1
        End If
        rst.MoveNext
    Loop
    
    rst.Close
    
    ' Zum Abschluß alle disqualifizierten Paare auf den letzten Platz
    ' in dieser Runde setzen
    Set rst = dbs.OpenRecordset(stmt)
    Do While (Not rst.EOF)
        If (rst!DQ_ID > 0) Then
            rst.Edit
            rst(wertungsrichter) = anzPaareOhneDQ + 1
            rst.Update
        End If
        
        rst.MoveNext
    Loop
    
    rst.Close

End Sub

Public Function fill_Paare_rundenquali(ziel, quelle, rt)
    ' überzählige löschen
    Dim sqlcmd As String
    sqlcmd = "delete from Paare_Rundenqualifikation pr where pr.rt_id=" & rt
    sqlcmd = sqlcmd & " and not exists (select 1 from Paare p where pr.tp_id=p.tp_id and p.anwesent_status>0)"
    DBEngine(0)(0).Execute (sqlcmd)
    ' neue hinzufügen
    If quelle.RecordCount > 0 Then quelle.MoveFirst
    Do Until quelle.EOF()
        ziel.AddNew
        ziel!TP_ID = quelle!TP_ID
        ziel!RT_ID = rt
        ziel!Anwesend_Status = quelle!Anwesend_Status
        ziel!Verein_Name = quelle!Verein_Name
        ziel!Rundennummer = Null
        ziel.Update
        quelle.MoveNext
    Loop
End Function

