Option Strict On
Option Explicit On

Imports ZoppaReports.Tokens

Namespace Tokens

    ''' <summary>トークン位置情報です。</summary>
    Public Structure TokenPosition

        ' 参照しているトークン
        Private ReadOnly mToken As IToken

        ''' <summary>トークン位置。</summary>
        Public ReadOnly Position As Integer

        ''' <summary>トークン型を取得する。</summary>
        ''' <returns>トークン型。</returns>
        Public ReadOnly Property TokenType As Type
            Get
                Return If(Me.mToken?.TokenType, Nothing)
            End Get
        End Property

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="src">参照するトークン。</param>
        ''' <param name="pos">トークン位置。</param>
        Public Sub New(src As IToken, pos As Integer)
            Me.mToken = src
            Me.Position = pos
        End Sub

        ''' <summary>型を指定してトークンを取得します。</summary>
        ''' <typeparam name="T">取得する型。</typeparam>
        ''' <returns>トークン情報。</returns>
        Friend Function GetToken(Of T As {Class, IToken})() As T
            Return TryCast(Me.mToken, T)
        End Function

        ''' <summary>文字列表現を取得する。</summary>
        ''' <returns>文字列。</returns>
        Public Overrides Function ToString() As String
            Return Me.mToken.ToString()
        End Function

    End Structure

End Namespace
