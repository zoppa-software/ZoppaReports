Imports Microsoft.Extensions.Logging
Imports Microsoft.VisualBasic.ApplicationServices
Imports ZoppaReports
Imports ZoppaLoggingExtensions

Namespace My
    ' 次のイベントは MyApplication に対して利用できます:
    ' Startup:アプリケーションが開始されたとき、スタートアップ フォームが作成される前に発生します。
    ' Shutdown:アプリケーション フォームがすべて閉じられた後に発生します。このイベントは、アプリケーションが異常終了したときには発生しません。
    ' UnhandledException:ハンドルされない例外がアプリケーションで発生したときに発生します。
    ' StartupNextInstance:単一インスタンス アプリケーションが起動され、それが既にアクティブであるときに発生します。 
    ' NetworkAvailabilityChanged:ネットワーク接続が接続されたとき、または切断されたときに発生します。
    Partial Friend Class MyApplication

        Private mLogFactory As ILoggerFactory

        Protected Overrides Function OnStartup(eventArgs As StartupEventArgs) As Boolean
            mLogFactory = LoggerFactory.Create(
                Sub(builder)
                    builder.AddZoppaLogging(
                        Sub(config)
                            config.MinimumLogLevel = LogLevel.Trace
                            config.DefaultLogFile = "report_test.log"
                        End Sub
                    )
                    builder.SetMinimumLevel(LogLevel.Trace)
                End Sub
             )
            SetZoppaReportsLogFactory(mLogFactory)

            Return MyBase.OnStartup(eventArgs)
        End Function

        Protected Overrides Sub OnShutdown()
            MyBase.OnShutdown()
            mLogFactory?.Dispose()
        End Sub

    End Class

End Namespace
