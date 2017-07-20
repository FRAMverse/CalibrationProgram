Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Collections.Generic
Imports System.IO
Module CheckCNRMod
    Sub CheckCNR()
        'Subroutine to check CNR data.
        '**************************************************************************

        'Print("Checking CNR data")

        'COMPUTE TOTAL CATCH IN EACH FISHERY by summing over stocks and ages
        'Report to errorfile fisheries without catch but CNR 
        'For TStep = 1 To NumSteps
        '    For Fish = 1 To NumFish
        '        For STk = MinStk To NumStk
        '            For Age = 2 To MaxAge
        '                TotCatch(Fish, TStep) = TotCatch(Fish, TStep) + Catcha(STk, Age, Fish, TStep)
        '            Next Age
        '        Next STk
        '    Next Fish
        'Next TStep
        Dim a As CWTData
        ReDim TotCatch(NumFish, NumSteps)
        For Each a In CWTList
            Catchb = a.cCatch
            TStep = a.cTStep
            Fish = a.cFish
            STk = a.cStk
            Age = a.cAge

            TotCatch(Fish, TStep) = TotCatch(Fish, TStep) + Catchb
        Next


        'For TStep = 1 To NumSteps
        '    For Fish = 1 To NumFish
        '        If TotCatch(Fish, TStep) = 0 And (CNRFlag(Fish, TStep) = 1 Or CNRFlag(Fish, TStep) = 2) Then
        '            Print(3, vbCrLf)
        '            Print(3, "CNR cannot be estimated for fishery with no catch", vbCrLf)
        '            Print(3, "Fish=", Fish, "TStep=", TStep, vbCrLf)
        '        End If
        '    Next Fish
        'Next TStep

        'For TStep = 1 To NumSteps
        '    For Fish = 1 To NumFish - 1
        '        If TotCatch(Fish, TStep) = 0 And CNRFlag(BaseType, Fish, TStep) > 0 Then
        '            Print(12, vbCrLf)
        '            Print(12, "CNR cannot be estimated for fishery with no catch.", vbCrLf)
        '            Print(12, "Fish=", Fish, "TStep=", TStep, vbCrLf)
        '            MsgBox("Fishery " & Fish & "TimeStep " & TStep & " has CNR but no landed catch. Please designate a surrogate fishery or time step for imputing recoveries.")
        '        End If
        '    Next
        'Next

    End Sub
End Module
