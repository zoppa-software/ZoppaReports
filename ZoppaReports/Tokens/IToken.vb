Option Strict On
Option Explicit On

Namespace Tokens

    ''' <summary>トークンインターフェイス。</summary>
    Public Interface IToken

        ''' <summary>格納されている値を取得する。</summary>
        ''' <returns>格納値。</returns>
        ReadOnly Property Contents() As Object

        ''' <summary>トークン型を取得する。</summary>
        ''' <returns>トークン型。</returns>
        ReadOnly Property TokenType As Type

    End Interface

End Namespace