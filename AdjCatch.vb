Module AdjCatchMod
    Sub adjcatch()
        '**********************************************************************
        'Subroutine to adjust recoveries in each fishery based upon the value
        ' of CatchFlag%:
        '    0 = No adjustment;
        '    1 = Adjust model catch to equal estimated catch;
        '    2 = Adjust model catch to equal estimated catch if model catch
        '        exceeds estimated catch.
        '    3 = use external adjustment
        '**********************************************************************

        'COMPUTE ADJUSTMENT FACTOR FOR EACH FISHERY
        ReDim RecAdjFactor(NumFish, NumSteps)
        ReDim FishExpCWTAll(NumStk, MaxAge, NumFish, NumSteps)

        Dim counter As Integer = 0
        Dim fishstring As String = " "
        Dim fishstringall As String = " "

        Dim CWTCatch As String
        CWTCatch = filepath & "CWTCatch"
        FileOpen(25, CWTCatch, OpenMode.Output)

        Dim Unadjusted As String
        Unadjusted = filepath & "\Unadjusted.txt"
        FileOpen(16, Unadjusted, OpenMode.Output)

        Print(16, "Fishery" & "TimeStep" & "," & "AnnualCatch" & "," & "TrueCatch" & "," & vbCrLf)




        For Fish = 1 To NumFish
            
            For TStep = 1 To NumSteps

                Print(16, Fish & "," & TStep & "," & AnnualCatch(Fish, TStep) & "," & TrueCatch(Fish, TStep) & vbCrLf)

                If Firstpass = True Then
                    CatchFlag(Fish, TStep) = 0
                End If

                If NoExpansions = True Then
                    CatchFlag(Fish, TStep) = 0
                End If

                'If AnnualCatch(Fish, TStep) > 0 And TrueCatch(Fish, TStep = 0) Then
                '    MsgBox("Catch in Fishery " & Fish & " and TStep " & TStep & " is zero. If you want to have BPERs, please enter a small, nominal catch.")
                'End If

                Select Case CatchFlag(Fish, TStep) 'located in cal file to the right of base period fishery catches or in BasePeriodCatch table of Calibration Support db

                    ' ADJUST MODEL CATCH TO ESTIMATED CATCH
                    Case 1
                        If AnnualCatch(Fish, TStep) > 0 Then
                            RecAdjFactor(Fish, TStep) = TrueCatch(Fish, TStep) / AnnualCatch(Fish, TStep) 'true catch = base period catch 
                        Else
                            If TrueCatch(Fish, TStep) > 0 Then
                                counter = counter + 1
                                fishstring = CStr(Fish) & " TStep " & CStr(TStep)
                                fishstringall = fishstringall & ", " & fishstring
                                'MsgBox("Base period catch exists yet recoveries are zero for fishery " & Fish)
                            End If
                        End If
                        ' For TStep = 1 To NumSteps
                        For STk = MinStk To NumStk
                            For Age = 2 To MaxAge
                                If STk = 9 And Age = 3 And Fish = 52 And TStep = 3 Then
                                    TStep = 3
                                End If
                                If CWTAll(STk, Age, Fish, TStep) > 0 Then
                                    FishExpCWTAll(STk, Age, Fish, TStep) = CWTAll(STk, Age, Fish, TStep) * RecAdjFactor(Fish, TStep)
                                End If
                            Next Age
                        Next STk
                        'Next TStep

                        ' ADJUST MODEL CATCH IF GREATER THAN ESTIMATED CATCH

                    Case 2
                        If AnnualCatch(Fish, TStep) > 0 Then
                            RecAdjFactor(Fish, TStep) = TrueCatch(Fish, TStep) / AnnualCatch(Fish, TStep)
                        Else
                            If TrueCatch(Fish, TStep) > 0 Then
                                counter = counter + 1
                                fishstring = CStr(Fish) & " TStep " & CStr(TStep)
                                fishstringall = fishstringall & ", " & fishstring
                                'MsgBox("Base period catch exists yet recoveries are zero for fishery " & Fish & " " & TStep)
                            End If
                        End If
                        If RecAdjFactor(Fish, TStep) < 1 Then
                            '  For TStep = 1 To NumSteps
                            For STk = MinStk To NumStk
                                For Age = 2 To MaxAge
                                    If CWTAll(STk, Age, Fish, TStep) > 0 Then
                                        FishExpCWTAll(STk, Age, Fish, TStep) = CWTAll(STk, Age, Fish, TStep) * RecAdjFactor(Fish, TStep)
                                    End If
                                Next Age
                            Next STk
                            ' Next TStep
                        Else
                            RecAdjFactor(Fish, TStep) = 99

                            '  For TStep = 1 To NumSteps
                            For STk = MinStk To NumStk
                                For Age = 2 To MaxAge
                                    If CWTAll(STk, Age, Fish, TStep) > 0 Then
                                        FishExpCWTAll(STk, Age, Fish, TStep) = CWTAll(STk, Age, Fish, TStep)
                                    End If
                                Next Age
                            Next STk
                            ' Next TStep
                        End If

                        ' ADJUST MODEL CATCH TO ESTIMATE CATCH


                    Case 3
                        If AnnualCatch(Fish, TStep) > 0 Then
                            RecAdjFactor(Fish, TStep) = TrueCatch(Fish, TStep) * ExternalModelStockProportion(Fish, TStep) / AnnualCatch(Fish, TStep) 'true catch = base period catch 
                        Else
                            If TrueCatch(Fish, TStep) > 0 Then
                                counter = counter + 1
                                fishstring = CStr(Fish) & " TStep " & CStr(TStep)
                                fishstringall = fishstringall & ", " & fishstring
                                'MsgBox("Base period catch exists yet recoveries are zero for fishery " & Fish)
                            End If
                        End If
                        ' For TStep = 1 To NumSteps
                        For STk = MinStk To NumStk
                            For Age = 2 To MaxAge
                                If CWTAll(STk, Age, Fish, TStep) > 0 Then
                                    FishExpCWTAll(STk, Age, Fish, TStep) = CWTAll(STk, Age, Fish, TStep) * RecAdjFactor(Fish, TStep)
                                End If
                            Next Age
                        Next STk
                        '  Next TStep


                        'NO ADJUSTMENT

                    Case Else
                        '   For TStep = 1 To NumSteps
                        For STk = MinStk To NumStk
                            For Age = 2 To MaxAge
                                If CWTAll(STk, Age, Fish, TStep) > 0 Then
                                    FishExpCWTAll(STk, Age, Fish, TStep) = CWTAll(STk, Age, Fish, TStep)
                                End If
                            Next Age
                        Next STk
                        '   Next TStep
                        RecAdjFactor(Fish, TStep) = 99
                End Select

                RecAdjFactor(Fish, TStep) = TrueCatch(Fish, TStep) / AnnualCatch(Fish, TStep)
                '****AHB 1/11/23 to alert user that a fishery/time step has zero catch, but CWTs are present
                If RecAdjFactor(Fish, TStep) = 0 And AnnualCatch(Fish, TStep) <> 0 Then
                    MsgBox("Catch in Fishery " & Fish & " and TStep " & TStep & " is zero. If you want to have BPERs, please enter a small, nominal catch in the BasePeriodCatch table.")
                End If
            Next TStep
        Next Fish
        FileClose(16)

        AdjustedCatch = True

        If counter > 0 Then
            MsgBox("Base period catch exists yet recoveries are zero for fishery " & fishstringall)
        End If
        If Firstpass <> True Then
            Print(25, "Stock,  Age, Fish, TStep, CWT" & vbCrLf)
            For STk = 1 To NumStk
                Print(25, "ESCExpansionFactor " & "," & STk & "," & EscExpFact(STk) & vbCrLf)
                For Age = 2 To MaxAge
                    For Fish = 1 To NumFish
                        For TStep = 1 To NumSteps
                            If CWTAll(STk, Age, Fish, TStep) > 0 Then
                                Print(25, STk & "," & Age & "," & Fish & "," & TStep & "," & FishExpCWTAll(STk, Age, Fish, TStep) & vbCrLf)
                            End If
                        Next
                    Next
                Next
            Next
        End If

    End Sub
End Module
