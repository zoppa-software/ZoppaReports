Option Strict On
Option Explicit On

Imports System.Drawing.Printing
Imports System.Windows.Controls

Namespace Designer

    ''' <summary>帳票サイズ。</summary>
    Public NotInheritable Class ReportsSize

        ''' <summary>ページ種類を取得する。</summary>
        Public ReadOnly Property Kind As PaperKind

        ''' <summary>ページ幅を取得する。</summary>
        Public ReadOnly Property WidthInMM As Double

        ''' <summary>ページ幅を取得する（インチ単位）</summary>
        Public ReadOnly Property WidthInInc As Double
            Get
                Return (WidthInMM / 0.254)
            End Get
        End Property

        ''' <summary>ページ高さを取得する。</summary>
        Public ReadOnly Property HeightInMM As Double

        ''' <summary>ページ高さを取得する。</summary>
        Public ReadOnly Property HeightInInc As Double
            Get
                Return (HeightInMM / 0.254)
            End Get
        End Property

        ''' <summary>ページ向きを取得する（インチ単位）</summary>
        Public ReadOnly Property Orientation As ReportsOrientation = ReportsOrientation.Portrait

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="kd">種類。</param>
        ''' <param name="wmm">幅。</param>
        ''' <param name="hmm">高さ。</param>
        ''' <param name="ort">向き。</param>
        Friend Sub New(kd As PaperKind, wmm As Double, hmm As Double, ort As ReportsOrientation)
            Me.Kind = kd
            Me.WidthInMM = wmm
            Me.HeightInMM = hmm
            Me.Orientation = ort
        End Sub

        ''' <summary>ページサイズを検索する。</summary>
        ''' <param name="name">ページ名。</param>
        ''' <returns>ページサイズ。</returns>
        Friend Shared Function Search(name As String) As ReportsSize
            Select Case name.ToUpper()
                Case "A2"
                    Return A2
                Case "A3"
                    Return A3
                Case "A4"
                    Return A4
                Case "A5"
                    Return A5
                Case "B4"
                    Return B4
                Case "B5"
                    Return B5
                Case Else
                    Throw New IndexOutOfRangeException("有効なページサイズが指定されていません")
            End Select
        End Function

        ''' <summary>A4サイズを取得する。</summary>
        Public Shared ReadOnly Property A2 As ReportsSize = New ReportsSize(PaperKind.A2, 420, 594, ReportsOrientation.Portrait)

        ''' <summary>A3サイズを取得する。</summary>
        Public Shared ReadOnly Property A3 As ReportsSize = New ReportsSize(PaperKind.A3, 297, 420, ReportsOrientation.Portrait)

        ''' <summary>A4サイズを取得する。</summary>
        Public Shared ReadOnly Property A4 As ReportsSize = New ReportsSize(PaperKind.A4, 210, 297, ReportsOrientation.Portrait)

        ''' <summary>A5サイズを取得する。</summary>
        Public Shared ReadOnly Property A5 As ReportsSize = New ReportsSize(PaperKind.A5, 148, 210, ReportsOrientation.Portrait)

        ''' <summary>A6サイズを取得する。</summary>
        Public Shared ReadOnly Property B4 As ReportsSize = New ReportsSize(PaperKind.B4, 257, 364, ReportsOrientation.Portrait)

        ''' <summary>B5サイズを取得する。</summary>
        Public Shared ReadOnly Property B5 As ReportsSize = New ReportsSize(PaperKind.B5, 182, 257, ReportsOrientation.Portrait)

        ''' <summary>文字列に変換する。</summary>
        ''' <param name="size">サイズ。</param>
        ''' <returns>文字列。</returns>
        Public Shared Widening Operator CType(size As ReportsSize) As String
            Return size.Kind.ToString()
        End Operator

    End Class

End Namespace
