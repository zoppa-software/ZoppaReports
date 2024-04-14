Option Strict On
Option Explicit On

Imports System.Drawing.Printing

Namespace Settings

    ''' <summary>帳票座標。</summary>
    Public Structure ReportsCoord

        ''' <summary>単位。</summary>
        Public Enum Unit As Byte

            ''' <summary>ピクセル。</summary>
            Pixel

            ''' <summary>インチ。</summary>
            Inch

            ''' <summary>ミリメートル。</summary>
            Millimeter

        End Enum

        Private ReadOnly mValue As Double

        Public ReadOnly Property Pixel As Double
            Get
                Return Me.mValue
            End Get
        End Property

        Public ReadOnly Property Inch As Double
            Get
                Return Me.mValue / 96
            End Get
        End Property

        Public ReadOnly Property Millimeter As Double
            Get
                Return Me.mValue / 96 * 25.4
            End Get
        End Property

        Private Sub New(value As Double)
            Me.mValue = value
        End Sub


        Public Shared Function Create(value As Double, unit As Unit) As ReportsCoord
            Select Case unit
                Case Unit.Pixel
                    Return New ReportsCoord(value)
                Case Unit.Inch
                    Return New ReportsCoord(value * 96)
                Case Unit.Millimeter
                    Return New ReportsCoord(value * 96 / 25.4)
                Case Else
                    Throw New ArgumentOutOfRangeException("無効な単位です。")
            End Select
        End Function

        Public Shared Function Zero() As ReportsCoord
            Return New ReportsCoord(0)
        End Function

        ''' <summary>文字列からReportsCoordに変換する。</summary>
        ''' <param name="inp">入力文字列。</param>
        ''' <returns>帳票座標。</returns>
        Public Shared Widening Operator CType(inp As String) As ReportsCoord
            inp = inp.ToLower()
            If inp.EndsWith("px") Then
                Dim ov = ParseValue(inp.Substring(0, inp.Length - 2))
                Return Create(ov, Unit.Pixel)
            ElseIf inp.EndsWith("mm") Then
                Dim ov = ParseValue(inp.Substring(0, inp.Length - 2))
                Return Create(ov, Unit.Millimeter)
            ElseIf inp.EndsWith("in") Then
                Dim ov = ParseValue(inp.Substring(0, inp.Length - 2))
                Return Create(ov, Unit.Inch)
            Else
                Dim ov = ParseValue(inp)
                Return Create(ov, Unit.Pixel)
            End If
        End Operator

        Private Shared Function ParseValue(value As String) As Double
            Dim outValue As Double
            If Double.TryParse(value, outValue) Then
                Return outValue
            Else
                Throw New InvalidCastException("文字列を ReportsCoordに変換できません")
            End If
        End Function

        ''' <summary>文字列に変換する。</summary>
        Public Overrides Function ToString() As String
            Return $"{Me.mValue}px"
        End Function

    End Structure

End Namespace
