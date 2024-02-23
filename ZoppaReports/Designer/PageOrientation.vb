Option Strict On
Option Explicit On

Imports System.Drawing.Printing

Namespace Designer

    ''' <summary>帳票の向き。</summary>
    Public NotInheritable Class PageOrientation

        ''' <summary>横向き。</summary>
        Public Shared ReadOnly Property Portrait As PageOrientation = New PageOrientation("portrait")

        ''' <summary>縦向き。</summary>
        Public Shared ReadOnly Property Landscape As PageOrientation = New PageOrientation("landscape")

        Private ReadOnly mValue As String

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="value">値。</param>
        Private Sub New(value As String)
            Me.mValue = value
        End Sub

        ''' <summary>文字列に変換する。</summary>
        ''' <param name="ori">向き。</param>
        ''' <returns>文字列。</returns>
        Public Shared Widening Operator CType(ori As PageOrientation) As String
            Return ori.mValue
        End Operator

        ''' <summary>文字列からPageOrientationに変換する</summary>
        ''' <param name="inp">入力文字列。</param>
        ''' <returns>帳票の向き。</returns>
        Public Shared Widening Operator CType(inp As String) As PageOrientation
            For Each ori In New PageOrientation() {Portrait, Landscape}
                If ori.mValue = inp Then
                    Return ori
                End If
            Next
            Throw New InvalidCastException("文字列を PageOrientationに変換できません")
        End Operator

        ''' <summary>等価演算子のオーバーロード。</summary>
        ''' <param name="lf">左辺。</param>
        ''' <param name="rt">右辺。</param>
        ''' <returns>等しければ真。</returns>
        Public Shared Operator =(ByVal lf As PageOrientation, ByVal rt As PageOrientation) As Boolean
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
        Public Shared Operator <>(ByVal lf As PageOrientation, ByVal rt As PageOrientation) As Boolean
            Return Not (lf = rt)
        End Operator

    End Class

End Namespace
