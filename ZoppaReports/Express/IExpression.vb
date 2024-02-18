Option Strict On
Option Explicit On

Imports Microsoft.Extensions.Logging
Imports ZoppaReports.Tokens

Namespace Express

    ''' <summary>式インターフェイス。</summary>
    Public Interface IExpression

        ''' <summary>式を実行する。</summary>
        ''' <param name="env">環境値情報。</param>
        ''' <returns>実行結果。</returns>
        Function Executes(env As Environments) As IToken

        ''' <summary>ログ出力を行います。</summary>
        ''' <param name="logger">ロガー。</param>
        ''' <param name="nestLv">ネストレベル。</param>
        Sub Logging(logger As ILogger, nestLv As Integer)

    End Interface

End Namespace