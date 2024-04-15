Option Strict On
Option Explicit On

Imports System.Drawing.Printing
Imports System.Windows.Controls

Namespace Designer

    ''' <summary>帳票サイズ。</summary>
    Public NotInheritable Class ReportsSize_old

#Region "size properties"

        ''' <summary>遅延インスタンス生成プロパティ。</summary>
        Private Shared ReadOnly Property LazyA2() As New Lazy(Of ReportsSize_old)(Function() New ReportsSize_old(PaperKind.A2, 420, 594, ReportsOrientation_old.Portrait))

        ''' <summary>A2サイズを取得する。</summary>
        Public Shared ReadOnly Property A2 As ReportsSize_old
            Get
                Return LazyA2.Value
            End Get
        End Property

        ''' <summary>遅延インスタンス生成プロパティ。</summary>
        Private Shared ReadOnly Property LazyA3() As New Lazy(Of ReportsSize_old)(Function() New ReportsSize_old(PaperKind.A3, 297, 420, ReportsOrientation_old.Portrait))

        ''' <summary>A3サイズを取得する。</summary>
        Public Shared ReadOnly Property A3 As ReportsSize_old
            Get
                Return LazyA3.Value
            End Get
        End Property

        ''' <summary>遅延インスタンス生成プロパティ。</summary>
        Private Shared ReadOnly Property LazyA4() As New Lazy(Of ReportsSize_old)(Function() New ReportsSize_old(PaperKind.A4, 210, 297, ReportsOrientation_old.Portrait))

        ''' <summary>A4サイズを取得する。</summary>
        Public Shared ReadOnly Property A4 As ReportsSize_old
            Get
                Return LazyA4.Value
            End Get
        End Property

        ''' <summary>遅延インスタンス生成プロパティ。</summary>
        Private Shared ReadOnly Property LazyA5() As New Lazy(Of ReportsSize_old)(Function() New ReportsSize_old(PaperKind.A5, 148, 210, ReportsOrientation_old.Portrait))

        ''' <summary>A5サイズを取得する。</summary>
        Public Shared ReadOnly Property A5 As ReportsSize_old
            Get
                Return LazyA5.Value
            End Get
        End Property

        ''' <summary>遅延インスタンス生成プロパティ。</summary>
        Private Shared ReadOnly Property LazyB4() As New Lazy(Of ReportsSize_old)(Function() New ReportsSize_old(PaperKind.B4, 257, 364, ReportsOrientation_old.Portrait))

        ''' <summary>B4サイズを取得する。</summary>
        Public Shared ReadOnly Property B4 As ReportsSize_old
            Get
                Return LazyB4.Value
            End Get
        End Property

        ''' <summary>遅延インスタンス生成プロパティ。</summary>
        Private Shared ReadOnly Property LazyB5() As New Lazy(Of ReportsSize_old)(Function() New ReportsSize_old(PaperKind.B5, 182, 257, ReportsOrientation_old.Portrait))

        ''' <summary>B5サイズを取得する。</summary>
        Public Shared ReadOnly Property B5 As ReportsSize_old
            Get
                Return LazyB5.Value
            End Get
        End Property

#End Region

        ''' <summary>ページ種類を取得する。</summary>
        Public ReadOnly Property Kind As PaperKind

        ''' <summary>ページ幅を取得する。</summary>
        Public ReadOnly Property WidthInMM As Double

        '''' <summary>ページ幅を取得する（インチ単位）</summary>
        'Public ReadOnly Property WidthInInc As Double
        '    Get
        '        Return (WidthInMM / 0.254)
        '    End Get
        'End Property

        ''' <summary>ページ高さを取得する。</summary>
        Public ReadOnly Property HeightInMM As Double

        '''' <summary>ページ高さを取得する。</summary>
        'Public ReadOnly Property HeightInInc As Double
        '    Get
        '        Return (HeightInMM / 0.254)
        '    End Get
        'End Property

        ''' <summary>ページ向きを取得する（インチ単位）</summary>
        Public ReadOnly Property Orientation As ReportsOrientation_old = ReportsOrientation_old.Portrait

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="kd">種類。</param>
        ''' <param name="wmm">幅。</param>
        ''' <param name="hmm">高さ。</param>
        ''' <param name="ort">向き。</param>
        Friend Sub New(kd As PaperKind, wmm As Double, hmm As Double, ort As ReportsOrientation_old)
            Me.Kind = kd
            Me.WidthInMM = wmm
            Me.HeightInMM = hmm
            Me.Orientation = ort
        End Sub

        ''' <summary>ページサイズを検索する。</summary>
        ''' <param name="name">ページ名。</param>
        ''' <returns>ページサイズ。</returns>
        Friend Shared Function Search(name As String) As ReportsSize_old
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

        ''' <summary>文字列に変換する。</summary>
        ''' <param name="size">サイズ。</param>
        ''' <returns>文字列。</returns>
        Public Shared Widening Operator CType(size As ReportsSize_old) As String
            Return size.ToString()
        End Operator

        ''' <summary>文字列からPageOrientationに変換する</summary>
        ''' <param name="inp">入力文字列。</param>
        ''' <returns>帳票の向き。</returns>
        Public Shared Widening Operator CType(inp As String) As ReportsSize_old
            If inp.StartsWith("custom") Then
                Dim prms = inp.Split(","c).Select(Of String)(Function(s) s.Trim()).ToArray()
                Return New ReportsSize_old(PaperKind.Custom, Convert.ToDouble(prms(1)), Convert.ToDouble(prms(2)), CTypeDynamic(Of ReportsOrientation_old)(prms(3)))
            Else
                Return Search(inp)
            End If
        End Operator

        ''' <summary>インスタンスの文字列表現を取得する。</summary>
        ''' <returns>文字列表現。</returns>
        Public Overrides Function ToString() As String
            If Me.Kind = PaperKind.Custom Then
                Return $"custom,{Me.WidthInMM},{Me.HeightInMM},{Me.Orientation}"
            Else
                Return Me.Kind.ToString()
            End If
        End Function

    End Class

End Namespace
