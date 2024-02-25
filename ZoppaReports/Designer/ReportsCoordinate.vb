Option Strict On
Option Explicit On

Imports System.Drawing.Printing
Imports System.Runtime.CompilerServices
Imports System.Windows.Controls

Namespace Designer

    Public Module ReportsCoordinate

        ''' <summary>mm を 1/72 インチに変換する係数。</summary>
        Public Const UNIT_COEFF As Double = 1 / (25.4 / 72)

        Public Const UNIT_COEFF_INCH As Double = 1 / 0.254

        <Extension>
        Public Function MmToPrint(pos As Double) As Double
            Return pos * UNIT_COEFF
        End Function

        <Extension>
        Public Function MmToInch(pos As Double) As Double
            Return pos * UNIT_COEFF_INCH
        End Function




    End Module

End Namespace
