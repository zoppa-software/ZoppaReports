Option Strict On
Option Explicit On

Imports ZoppaReports.Tokens

Namespace ExTokens

    ''' <summary>Forトークン。</summary>
    Public NotInheritable Class ForToken
        Implements IToken

        ' 条件式トークン
        Private ReadOnly mToken As List(Of TokenPosition)

        ''' <summary>格納されている値を取得する。</summary>
        ''' <returns>格納値。</returns>
        Public ReadOnly Property Contents As Object Implements IToken.Contents
            Get
                Return Me.mToken
            End Get
        End Property

        ''' <summary>トークン型を取得する。</summary>
        ''' <returns>トークン型。</returns>
        Public ReadOnly Property TokenType As Type Implements IToken.TokenType
            Get
                Return GetType(ForToken)
            End Get
        End Property

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="tokens">ループ変数のトークン。</param>
        Public Sub New(tokens As List(Of TokenPosition))
            Me.mToken = tokens
        End Sub

        ''' <summary>文字列条件を取得します。</summary>
        ''' <returns>文字列表現。</returns>
        Public Overrides Function ToString() As String
            Return "For"
        End Function

    End Class

End Namespace