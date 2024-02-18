Option Strict On
Option Explicit On

Imports Microsoft.Extensions.Logging
Imports Microsoft.Extensions.DependencyInjection
Imports System.Runtime.CompilerServices
Imports ZoppaReports.Tokens

Partial Module ZoppaReports

    ' サービスプロバイダー
    Private _provider As IServiceProvider = Nothing

    ' ログファクトリ
    Private _logFactory As ILoggerFactory = Nothing

    ' ロガー
    Private ReadOnly _logger As New Lazy(Of ILogger)(
        Function()
            If _provider IsNot Nothing Then
                Return _provider.GetService(Of ILoggerFactory)()?.CreateLogger("ZoppaReports")
            ElseIf _logFactory IsNot Nothing Then
                Return _logFactory.CreateLogger("ZoppaReports")
            Else
                Return Nothing
            End If
        End Function
    )

    ''' <summary>ログ出力オブジェクトを取得します。</summary>
    ''' <returns>ログ出力オブジェクト。</returns>
    Friend ReadOnly Property Logger As Lazy(Of ILogger)
        Get
            Return _logger
        End Get
    End Property

    ''' <summary>外部サービスプロバイダーからログ出力を行います。</summary>
    ''' <param name="provider">サービスプロバイダー。</param>
    ''' <returns>サービスプロバイダー。</returns>
    <Extension>
    Public Function SetZoppaReportsLogProvider(provider As IServiceProvider) As IServiceProvider
        If _provider Is Nothing Then
            _provider = provider
            Using scope = _logger.Value?.BeginScope("log setting")
                _logger.Value?.LogInformation("use other log by service provider")
            End Using
        End If
        Return _provider
    End Function

    ''' <summary>外部ログファクトリからログ出力を行います。</summary>
    ''' <param name="factory">ログファクトリ。</param>
    ''' <returns>ログファクトリ。</returns>
    <Extension>
    Public Function SetZoppaReportsLogFactory(factory As ILoggerFactory) As ILoggerFactory
        If _logFactory Is Nothing Then
            _logFactory = factory
            Using scope = _logger.Value?.BeginScope("log setting")
                _logger.Value?.LogInformation("use other log by logger factory")
            End Using
        End If
        Return _logFactory
    End Function

    ''' <summary>トークンリストをログ出力します。</summary>
    ''' <param name="tokens">トークンリスト。</param>
    Private Sub LoggingTokens(tokens As IEnumerable(Of TokenPosition))
        Logger.Value?.LogTrace("token count : {tokens}", tokens.Count)
        If Logger.Value?.IsEnabled(LogLevel.Trace) Then
            For Each tkn In tokens
                Logger.Value?.LogTrace("・{tkn} ({tkn.Position})", tkn, tkn.Position)
            Next
        End If
    End Sub

End Module
