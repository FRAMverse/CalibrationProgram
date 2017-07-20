Module CompEscapeMod
    Sub CompEscape()
        ReDim Escape(MaxAge, NumSteps)


        For Each CWTRow In CWTList
            If CWTRow.cLookUp = DictionaryKey Then
                Age = CWTRow.cAge
                Fish = CWTRow.cFish
                BY = CWTRow.cBY
                STk = CWTRow.cStk
                Stage = CWTRow.cStage
                TStep = CWTRow.cTStep
                ' make expansion based on ETRS not just escapement
                If Fish = 72 Or Fish = 73 Or Fish = 74 Then
                    Escape(Age, TStep) = Escape(Age, TStep) + CWTRow.cCatch
                End If

                'If Fish = 74 Then
                '    Escape(Age, TStep) = Escape(Age, TStep) + CWTRow.cCatch
                'End If
            End If
        Next
      
    End Sub
End Module
