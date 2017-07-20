Imports System.Math
Module NormlDistribution
    Function NormlDistr(ByVal X, ByVal Mean, ByVal SD)
        'Obtained from WDF/NBS program, Subroutine NORMFN.
        Z = (X - Mean) / SD
        AbsZ = Math.Abs(Z)
        Area = 1 / (1 + AbsZ * (AbsZ * (AbsZ * (AbsZ * (AbsZ * (0.000005383# * AbsZ + 0.0000488906#) + 0.0000380036#) + 0.0032776263#) + 0.0211410061#) + 0.049867347#))
        Area = 1 - 0.5 * Area ^ 16
        If Z < 0 Then
            Area = 1 - Area
        End If
        NormlDistr = Area
    End Function
End Module
