Option Strict On
Option Explicit On
Imports System.Text

Namespace Tokens

    ''' <summary>配列トークン。</summary>
    Public NotInheritable Class ArrayToken
        Implements IToken

        ' 値
        Private ReadOnly mValue As IToken()

        ''' <summary>格納されている値を取得する。</summary>
        ''' <returns>格納値。</returns>
        Public ReadOnly Property Contents As Object Implements IToken.Contents
            Get
                Return Me.mValue.Select(Of Object)(Function(x) x.Contents).ToArray()
            End Get
        End Property

        ''' <summary>トークン型を取得する。</summary>
        ''' <returns>トークン型。</returns>
        Public ReadOnly Property TokenType As Type Implements IToken.TokenType
            Get
                Return GetType(ArrayToken)
            End Get
        End Property

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="value">文字列。</param>
        Public Sub New(value As IToken())
            Me.mValue = value
        End Sub

        ''' <summary>文字列条件を取得します。</summary>
        ''' <returns>文字列表現。</returns>
        Public Overrides Function ToString() As String
            Dim res As New StringBuilder()
            For i As Integer = 0 To Me.mValue.Length - 1
                res.Append(Me.mValue(i).ToString())
                If i < Me.mValue.Length - 1 Then
                    res.Append(", ")
                End If
            Next
            Return res.ToString()
        End Function

    End Class

End Namespace
