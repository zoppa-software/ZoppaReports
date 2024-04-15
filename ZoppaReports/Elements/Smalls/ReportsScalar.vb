Option Strict On
Option Explicit On

Imports System.Drawing.Printing

Namespace Smalls

    ''' <summary>帳票座標。</summary>
    Public Structure ReportsScalar

        ''' <summary>単位。</summary>
        Public Enum Unit As Byte

            ''' <summary>ピクセル。</summary>
            Pixel

            ''' <summary>インチ。</summary>
            Inch

            ''' <summary>ミリメートル。</summary>
            Millimeter

        End Enum

        ' 値
        Private ReadOnly mValue As Double

#Region "properties"

        ''' <summary>ピクセルの値を取得します。</summary>
        Public ReadOnly Property Pixel As Double
            Get
                Return Me.mValue
            End Get
        End Property

        ''' <summary>インチの値を取得します。</summary>
        Public ReadOnly Property Inch As Double
            Get
                Return Me.mValue / 96
            End Get
        End Property

        ''' <summary>ミリメートルの値を取得します。</summary>
        Public ReadOnly Property Millimeter As Double
            Get
                Return Me.mValue / 96 * 25.4
            End Get
        End Property

#End Region

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="value">値。</param>
        Private Sub New(value As Double)
            Me.mValue = value
        End Sub

        ''' <summary>値と単位からReportsScalarを生成する。</summary>
        ''' <param name="value">値。</param>
        ''' <param name="unit">単位。</param>
        ''' <returns>ReportsScalar。</returns>
        Public Shared Function Create(value As Double, unit As Unit) As ReportsScalar
            Select Case unit
                Case Unit.Pixel
                    Return New ReportsScalar(value)
                Case Unit.Inch
                    Return New ReportsScalar(value * 96)
                Case Unit.Millimeter
                    Return New ReportsScalar(value * 96 / 25.4)
                Case Else
                    Throw New ArgumentOutOfRangeException("無効な単位です。")
            End Select
        End Function

        ''' <summary>ゼロを取得する。</summary>
        Public Shared Function Zero() As ReportsScalar
            Return New ReportsScalar(0)
        End Function

        ''' <summary>文字列からReportsCoordに変換する。</summary>
        ''' <param name="inp">入力文字列。</param>
        ''' <returns>帳票座標。</returns>
        Public Shared Widening Operator CType(inp As String) As ReportsScalar
            inp = inp.Trim().ToLower()
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

        ''' <summary>文字列を数値に変換する。</summary>
        ''' <param name="value">文字列。</param>
        ''' <returns>数値。</returns>
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
