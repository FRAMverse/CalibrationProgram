Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Collections.Generic
Imports System.IO
Module CheckLeg
    Sub CheckLegal()
        '**************************************************************************
        'Subroutine to check if legal population exists in for each recovery
        ' In the absence of a legal population, the data is written
        ' to an .ERR file.
        '**************************************************************************

        'Print("Checking for legal population")

        Print(12, "Cases with catch but legal population = 0. These recoveries are eliminated from analysis", vbCrLf)
        Print(12, "STK", "AGE", "FISH", "TIME", "Stage", "BY  ", "CATCH", vbCrLf)
        'For TStep = 1 To NumSteps
        '    For Fish = 1 To NumFish
        '        For STk = MinStk To NumStk
        '            For Age = 2 To MaxAge
        '                If Catcha(STk, Age, Fish, TStep) > 0 Then
        '                    TermStat = TermFlag(Fish, TStep)
        '                    CompLegProp()
        '                    If LegalProp <= 0 Then
        '                        Print(3, STk, Age, Fish, TStep, Catcha(STk, Age, Fish, TStep).ToString("0.00"), vbCrLf)
        '                    Else

        '                        Print(5, "1".PadLeft(4))
        '                        Print(5, STk.ToString.PadLeft(4))
        '                        Print(5, Age.ToString.PadLeft(4))
        '                        Print(5, Fish.ToString.PadLeft(4))
        '                        Print(5, TStep.ToString.PadLeft(4))
        '                        Print(5, Catcha(STk, Age, Fish, TStep).ToString("0.0000").PadLeft(12))
        '                        Print(5, vbCrLf)
        '                    End If
        '                End If
        '            Next Age
        '        Next STk
        '    Next Fish
        'Next TStep
       


        If Firstpass = True Then
            Dim a As CWTData
            Dim Errlist As New List(Of CWTData)
            ReDim Indextracker(CWTList.Count)
            Dim Elementcounter As Integer = -1
            For Each a In CWTList
                Elementcounter = Elementcounter + 1
                If a.cCatch > 0 Then
                    Fish = a.cFish
                    FishNum = a.cFish
                    TStep = a.cTStep
                    STk = a.cStk
                    Stage = a.cStage
                    Age = a.cAge
                    BY = a.cBY
                    FishYear = BY + Age
                    TermStat = TermFlag(Fish, TStep)
                    CompLegProp()
                    If LegalProp <= 0 Then
                        Print(12, STk, Age, Fish, TStep, Stage, BY, a.cCatch.ToString("0.00"), vbCrLf)
                        'Errlist.Add(a)
                        Indextracker(Elementcounter) = 1
                    End If
                End If
            Next
            'remove recoveries with zero legal population
            Dim i As Integer
            Dim b As Integer = 0
            For i = 0 To UBound(Indextracker) - 1
                If Indextracker(i) = 1 Then
                    CWTList.RemoveAt(i + b)
                    b = b - 1
                End If
            Next


        Else
            Dim counter As Integer = 0
            For STk = 1 To NumStk
                For Age = 2 To MaxAge
                    For Fish = 1 To NumFish
                        FishNum = Fish
                        For TStep = 1 To NumSteps
                            If STk = 1 And Age = 2 And Fish = 11 And TStep = 2 Then
                                TStep = 2
                            End If

                            If CWTAll(STk, Age, Fish, TStep) > 0 Then
                                
                                TermStat = TermFlag(Fish, TStep)
                                FishYear = BaseYear
                                CompLegProp()
                                If LegalProp <= 0 Then
                                    counter = counter + 1
                                    Print(12, STk, Age, Fish, TStep, "Stage", BaseYear, CWTAll(STk, Age, Fish, TStep), vbCrLf)
                                    CWTAll(STk, Age, Fish, TStep) = 0
                                Else
                                End If
                                If CWTAll(STk, Age, Fish, TStep) > 0 Then
                                    Print(11, STk & "," & Age & "," & Fish & "," & TStep & "," & CWTAll(STk, Age, Fish, TStep) & vbCrLf)
                                End If
                            End If
                        Next
                    Next
                Next
            Next
            If counter > 0 Then
                MsgBox(" There were " & counter & " instances of recoveries for stocks without a legal population size. Please see Errfile for more information.")
            End If

        End If
        FileClose(11)


    End Sub
End Module
