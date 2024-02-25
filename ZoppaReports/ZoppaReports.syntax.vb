Option Strict On
Option Explicit On

Imports System.Runtime.CompilerServices
Imports Microsoft.Extensions.Logging
Imports ZoppaReports.Tokens
Imports ZoppaReports.Lexical
Imports ZoppaReports.Express
Imports ZoppaReports.Parser

Partial Module ZoppaReports

    ''' <summary>解析したSQLのキャッシュサイズ。</summary>
    Public Const MAX_HISTORY_SIZE As Integer = 64

    ' コマンド履歴
    Private ReadOnly _cmdHistory As New List(Of (String, List(Of TokenPosition)))(MAX_HISTORY_SIZE \ 2)

    <Extension>
    Public Function Executes(expression As String, Optional parameter As Object = Nothing) As IToken
        Using scope = _logger.Value?.BeginScope(NameOf(Executes))
            Try
                Logger.Value?.LogDebug("executes expression : {expression}", expression)
                Return ExecutesRefEnvironments(expression, New Environments(parameter))

            Catch ex As Exception
                _logger.Value?.LogError("message:{ex.Message} stack trace:{ex.StackTrace}", ex.Message, ex.StackTrace)
                Throw
            End Try
        End Using
    End Function

    Friend Function ExecutesRefEnvironments(expression As String, parameter As Environments) As IToken
        If Logger.Value IsNot Nothing Then
            parameter.Logging(Logger.Value)
        End If

        ' トークンリストを取得
        Dim tokens = GetNewOrHistory(expression, AddressOf LexicalAnalysis.SplitToken)
        LoggingTokens(tokens)

        ' 評価
        Dim res = ParserAnalysis.Executes(tokens, parameter)
        Logger.Value?.LogDebug("answer expression : {ans}", res.ans)
        If Logger.Value IsNot Nothing Then
            res.expr.Logging(Logger.Value, 0)
        End If

        Return res.ans
    End Function

    ''' <summary>あれば履歴からトークンリストを取得、なければ新規作成して履歴に追加します。</summary>
    ''' <param name="sqlQuery">評価する文字列。</param>
    ''' <returns>トークンリスト。</returns>
    Private Function GetNewOrHistory(sqlQuery As String, splitFunc As Func(Of String, List(Of TokenPosition))) As List(Of TokenPosition)
        SyncLock _cmdHistory
            ' 履歴にあればぞれを返す
            For i As Integer = 0 To _cmdHistory.Count - 1
                If _cmdHistory(i).Item1 = sqlQuery Then
                    Return _cmdHistory(i).Item2
                End If
            Next

            ' 履歴になければ新規作成
            Dim tokens = splitFunc(sqlQuery)
            _cmdHistory.Add((sqlQuery, tokens))

            If _cmdHistory.Count > MAX_HISTORY_SIZE Then
                _cmdHistory.RemoveAt(0)
            End If

            Return tokens
        End SyncLock
    End Function

End Module
