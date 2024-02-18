Option Strict On
Option Explicit On

Imports Microsoft.Extensions.Logging
Imports ZoppaReports.Exceptions
Imports ZoppaReports.Tokens

Namespace Express

    ''' <summary>括弧内部式。</summary>
    Public NotInheritable Class ParenExpress
        Implements IExpression

        ' 内部式
        Private ReadOnly mInner As IExpression

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="inner">内部式。</param>
        Public Sub New(inner As IExpression)
            Me.mInner = inner
        End Sub

        ''' <summary>式を実行する。</summary>
        ''' <param name="env">環境値情報。</param>
        ''' <returns>実行結果。</returns>
        Public Function Executes(env As Environments) As IToken Implements IExpression.Executes
            Dim ans = Me.mInner?.Executes(env)
            If ans IsNot Nothing Then
                Return ans
            Else
                Throw New ReportsAnalysisException("内部式が実行できません")
            End If
        End Function

        ''' <summary>ログ出力を行います。</summary>
        ''' <param name="logger">ロガー。</param>
        ''' <param name="nestLv">ネストレベル。</param>
        Public Sub Logging(logger As ILogger, nestLv As Integer) Implements IExpression.Logging
            logger?.LogTrace(New String(" "c, nestLv * 2) & Me.ToString())
            Me.mInner?.Logging(logger, nestLv + 1)
        End Sub

        ''' <summary>文字列条件を取得します。</summary>
        ''' <returns>文字列表現。</returns>
        Public Overrides Function ToString() As String
            Return Me.mInner.ToString()
        End Function

    End Class

End Namespace