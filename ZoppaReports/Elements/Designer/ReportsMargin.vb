Option Strict On
Option Explicit On

Imports System.Drawing.Printing
Imports System.Windows.Controls

Namespace Designer

    ''' <summary>帳票のマージン。</summary>
    Public Structure ReportsMargin

        ''' <summary>左マージンを設定、取得します。</summary>
        Public Property Left As Double

        ''' <summary>上マージンを設定、取得します。</summary>
        Public Property Top As Double

        ''' <summary>右マージンを設定、取得します。</summary>
        Public Property Right As Double

        ''' <summary>下マージンを設定、取得します。</summary>
        Public Property Bottom As Double

        ''' <summary>文字列に変換する。</summary>
        ''' <param name="size">サイズ。</param>
        ''' <returns>文字列。</returns>
        Public Shared Widening Operator CType(size As ReportsMargin) As String
            Return size.ToString()
        End Operator

        ''' <summary>文字列からPageOrientationに変換する</summary>
        ''' <param name="inp">入力文字列。</param>
        ''' <returns>帳票の向き。</returns>
        Public Shared Widening Operator CType(inp As String) As ReportsMargin
            Try
                Dim values = inp.Split(","c)
                Select Case values.Length
                    Case 1
                        Dim v = Double.Parse(values(0).Trim())
                        Return New ReportsMargin With {.Left = v, .Top = v, .Right = v, .Bottom = v}
                    Case 2
                        Dim v1 = Double.Parse(values(0).Trim())
                        Dim v2 = Double.Parse(values(1).Trim())
                        Return New ReportsMargin With {.Left = v1, .Top = v2, .Right = v1, .Bottom = v2}
                    Case 3
                        Dim v1 = Double.Parse(values(0).Trim())
                        Dim v2 = Double.Parse(values(1).Trim())
                        Dim v3 = Double.Parse(values(2).Trim())
                        Return New ReportsMargin With {.Left = v1, .Top = v2, .Right = v3, .Bottom = v2}
                    Case 4
                        Dim v1 = Double.Parse(values(0).Trim())
                        Dim v2 = Double.Parse(values(1).Trim())
                        Dim v3 = Double.Parse(values(2).Trim())
                        Dim v4 = Double.Parse(values(3).Trim())
                        Return New ReportsMargin With {.Left = v1, .Top = v2, .Right = v3, .Bottom = v4}
                    Case Else
                        Throw New InvalidCastException("文字列を ReportsMarginに変換できません")
                End Select
            Catch ex As Exception
                Throw New InvalidCastException("文字列を ReportsMarginに変換できません", ex)
            End Try
        End Operator

        ''' <summary>文字列条件を取得します。</summary>
        ''' <returns>文字列表現。</returns>
        Public Overrides Function ToString() As String
            Return $"{Me.Left},{Me.Top},{Me.Right},{Me.Bottom}"
        End Function

    End Structure

End Namespace
