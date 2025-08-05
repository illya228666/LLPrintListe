Public Function LLPrintListeSuper(frm As Form, _
                                  LL1 As ListLabel.ListLabel, _
                                  BelegID As Long, _
                                  Mode As Integer, _
                                  Optional tmp As Boolean, _
                                  Optional Save As Boolean, _
                                  Optional SammelDruck As Boolean, _
                                  Optional Bereich As Integer = 2800) As Long
    'Wrapper that dispatches printing to the correct module based on Bereich.
    Select Case Bereich
        Case 2800
            LLPrintListeSuper = LLPrintListe2800(frm, LL1, BelegID, Mode, tmp, Save, SammelDruck)
        Case 5640
            LLPrintListeSuper = LLPrintListe5640(frm, LL1, BelegID, Mode, tmp, Save, SammelDruck)
        Case Else
            LLPrintListeSuper = -1
    End Select
End Function
