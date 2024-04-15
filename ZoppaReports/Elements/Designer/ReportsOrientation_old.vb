Option Strict On
Option Explicit On

Imports System.Drawing.Printing

Namespace Designer

    ''' <summary>帳票の向き。</summary>
    Public NotInheritable Class ReportsOrientation_old

#Region "orientation properties"

        ''' <summary>遅延インスタンス生成プロパティ。</summary>
        Private Shared ReadOnly Property LazyPortrait() As New Lazy(Of ReportsOrientation_old)(Function() New ReportsOrientation_old("portrait"))

        ''' <summary>横向き。</summary>
        Public Shared ReadOnly Property Portrait As ReportsOrientation_old
            Get
                Return LazyPortrait.Value
            End Get
        End Property

        ''' <summary>遅延インスタンス生成プロパティ。</summary>
        Private Shared ReadOnly Property LazyLandscape() As New Lazy(Of ReportsOrientation_old)(Function() New ReportsOrientation_old("landscape"))

        ''' <summary>縦向き。</summary>
        Public Shared ReadOnly Property Landscape As ReportsOrientation_old
            Get
                Return LazyLandscape.Value
            End Get
        End Property

#End Region

        ''' <summary>値。</summary>
        Private ReadOnly mValue As String

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="value">値。</param>
        Private Sub New(value As String)
            Me.mValue = value
        End Sub

        ''' <summary>文字列に変換する。</summary>
        ''' <param name="ori">向き。</param>
        ''' <returns>文字列。</returns>
        Public Shared Widening Operator CType(ori As ReportsOrientation_old) As String
            Return ori.mValue
        End Operator

        ''' <summary>文字列からPageOrientationに変換する</summary>
        ''' <param name="inp">入力文字列。</param>
        ''' <returns>帳票の向き。</returns>
        Public Shared Widening Operator CType(inp As String) As ReportsOrientation_old
            For Each ori In New ReportsOrientation_old() {Portrait, Landscape}
                If ori.mValue = inp Then
                    Return ori
                End If
            Next
            Throw New InvalidCastException("文字列を PageOrientationに変換できません")
        End Operator

        ''' <summary>等しければ真を返す。</summary>
        ''' <param name="obj">比較対象。</param>
        ''' <returns>比較結果。</returns>
        Public Overrides Function Equals(obj As Object) As Boolean
            Dim other = TryCast(obj, ReportsOrientation_old)
            If other IsNot Nothing Then
                Return (Me.mValue = other.mValue)
            End If
            Return False
        End Function

        ''' <summary>ハッシュコード値を取得する。</summary>
        ''' <returns>ハッシュコード値。</returns>
        Public Overrides Function GetHashCode() As Integer
            Return Me.mValue.GetHashCode()
        End Function

        ''' <summary>インスタンスの文字列表現を取得する。</summary>
        ''' <returns>文字列表現。</returns>
        Public Overrides Function ToString() As String
            Return Me.mValue
        End Function

        ''' <summary>等価演算子のオーバーロード。</summary>
        ''' <param name="lf">左辺。</param>
        ''' <param name="rt">右辺。</param>
        ''' <returns>等しければ真。</returns>
        Public Shared Operator =(ByVal lf As ReportsOrientation_old, ByVal rt As ReportsOrientation_old) As Boolean
            If (lf IsNot Nothing) AndAlso (rt IsNot Nothing) Then
                Return (lf.mValue = rt.mValue)
            Else
                Return False
            End If
        End Operator

        ''' <summary>不等価演算子のオーバーロード。</summary>
        ''' <param name="lf">左辺。</param>
        ''' <param name="rt">右辺。</param>
        ''' <returns>等しければ真。</returns>
        Public Shared Operator <>(ByVal lf As ReportsOrientation_old, ByVal rt As ReportsOrientation_old) As Boolean
            Return Not (lf = rt)
        End Operator

    End Class

End Namespace
