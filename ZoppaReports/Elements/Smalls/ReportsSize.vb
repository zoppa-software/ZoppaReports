Option Strict On
Option Explicit On

Imports System.Drawing.Printing

Namespace Smalls

    ''' <summary>帳票サイズ。</summary>
    Public NotInheritable Class ReportsSize

        ' 幅
        Private ReadOnly mWidth As ReportsScalar = ReportsScalar.Zero

        ' 高さ
        Private ReadOnly mHeight As ReportsScalar = ReportsScalar.Zero

#Region "size properties"

        ''' <summary>ページ種類を取得する。</summary>
        Public ReadOnly Property Kind As PaperKind

        ''' <summary>遅延インスタンス生成プロパティ。</summary>
        Private Shared ReadOnly Property LazyA2() As New Lazy(Of ReportsSize)(Function() New ReportsSize(PaperKind.A2, 420, 594))

        ''' <summary>A2サイズを取得する。</summary>
        Public Shared ReadOnly Property A2 As ReportsSize
            Get
                Return LazyA2.Value
            End Get
        End Property

        ''' <summary>遅延インスタンス生成プロパティ。</summary>
        Private Shared ReadOnly Property LazyA3() As New Lazy(Of ReportsSize)(Function() New ReportsSize(PaperKind.A3, 297, 420))

        ''' <summary>A3サイズを取得する。</summary>
        Public Shared ReadOnly Property A3 As ReportsSize
            Get
                Return LazyA3.Value
            End Get
        End Property

        ''' <summary>遅延インスタンス生成プロパティ。</summary>
        Private Shared ReadOnly Property LazyA4() As New Lazy(Of ReportsSize)(Function() New ReportsSize(PaperKind.A4, 210, 297))

        ''' <summary>A4サイズを取得する。</summary>
        Public Shared ReadOnly Property A4 As ReportsSize
            Get
                Return LazyA4.Value
            End Get
        End Property

        ''' <summary>遅延インスタンス生成プロパティ。</summary>
        Private Shared ReadOnly Property LazyA5() As New Lazy(Of ReportsSize)(Function() New ReportsSize(PaperKind.A5, 148, 210))

        ''' <summary>A5サイズを取得する。</summary>
        Public Shared ReadOnly Property A5 As ReportsSize
            Get
                Return LazyA5.Value
            End Get
        End Property

        ''' <summary>遅延インスタンス生成プロパティ。</summary>
        Private Shared ReadOnly Property LazyB4() As New Lazy(Of ReportsSize)(Function() New ReportsSize(PaperKind.B4, 257, 364))

        ''' <summary>B4サイズを取得する。</summary>
        Public Shared ReadOnly Property B4 As ReportsSize
            Get
                Return LazyB4.Value
            End Get
        End Property

        ''' <summary>遅延インスタンス生成プロパティ。</summary>
        Private Shared ReadOnly Property LazyB5() As New Lazy(Of ReportsSize)(Function() New ReportsSize(PaperKind.B5, 182, 257))

        ''' <summary>B5サイズを取得する。</summary>
        Public Shared ReadOnly Property B5 As ReportsSize
            Get
                Return LazyB5.Value
            End Get
        End Property

        ''' <summary>幅を取得する。</summary>
        Public ReadOnly Property Width As ReportsScalar
            Get
                Return Me.mWidth
            End Get
        End Property

        ''' <summary>高さを取得する。</summary>
        Public ReadOnly Property Height As ReportsScalar
            Get
                Return Me.mHeight
            End Get
        End Property

#End Region

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="kd">種類。</param>
        ''' <param name="wmm">幅。</param>
        ''' <param name="hmm">高さ。</param>
        Private Sub New(kd As PaperKind, wmm As Double, hmm As Double)
            Me.Kind = kd
            Me.mWidth = ReportsScalar.Create(wmm, ReportsScalar.Unit.Millimeter)
            Me.mHeight = ReportsScalar.Create(hmm, ReportsScalar.Unit.Millimeter)
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

        ''' <summary>文字列に変換する。</summary>
        ''' <param name="size">サイズ。</param>
        ''' <returns>文字列。</returns>
        Public Shared Widening Operator CType(size As ReportsSize) As String
            Return size.ToString()
        End Operator

        ''' <summary>文字列からPageOrientationに変換する</summary>
        ''' <param name="inp">入力文字列。</param>
        ''' <returns>帳票の向き。</returns>
        Public Shared Widening Operator CType(inp As String) As ReportsSize
            If inp.StartsWith("custom") Then
                Dim prms = inp.Split(","c).Select(Of String)(Function(s) s.Trim()).ToArray()
                Return New ReportsSize(PaperKind.Custom, CType(prms(1), ReportsScalar).Millimeter, CType(prms(2), ReportsScalar).Millimeter)
            Else
                Return Search(inp)
            End If
        End Operator

        ''' <summary>インスタンスの文字列表現を取得する。</summary>
        ''' <returns>文字列表現。</returns>
        Public Overrides Function ToString() As String
            If Me.Kind = PaperKind.Custom Then
                Return $"custom,{Me.Width},{Me.Height}"
            Else
                Return Me.Kind.ToString()
            End If
        End Function

    End Class

End Namespace