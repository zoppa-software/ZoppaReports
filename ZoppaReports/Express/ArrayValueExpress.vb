Option Strict On
Option Explicit On

Imports System.Reflection
Imports Microsoft.Extensions.Logging
Imports ZoppaReports.Exceptions
Imports ZoppaReports.Tokens

Namespace Express

    ''' <summary>配列値式。</summary>
    Public NotInheritable Class ArrayValueExpress
        Implements IExpression

        ' 対象トークン
        Private ReadOnly mToken As IdentToken

        ' 添字トークン
        Private ReadOnly mIndexToken As ParenExpress

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="token">対象トークン。</param>
        ''' <param name="indexToken">インデックストークン。</param>
        Public Sub New(token As IdentToken, indexToken As ParenExpress)
            If token IsNot Nothing AndAlso indexToken IsNot Nothing Then
                Me.mToken = token
                Me.mIndexToken = indexToken
            Else
                Throw New ReportsAnalysisException("値式の生成にNullは使用できません")
            End If
        End Sub

        ''' <summary>式を実行する。</summary>
        ''' <param name="env">環境値情報。</param>
        ''' <returns>実行結果。</returns>
        Public Function Executes(env As Environments) As IToken Implements IExpression.Executes
            Try
                Dim arr = CType(env.GetValue(If(Me.mToken.Contents?.ToString(), "")), IList)
                Dim idx = Convert.ToInt32(Me.mIndexToken.Executes(env).Contents)
                Return ValueExpress.ConvertValueToken(arr(idx))
            Catch ex As Exception
                Throw New ReportsAnalysisException("配列の値を取得できませんでした", ex)
            End Try
        End Function

        ''' <summary>ログ出力を行います。</summary>
        ''' <param name="logger">ロガー。</param>
        ''' <param name="nestLv">ネストレベル。</param>
        Public Sub Logging(logger As ILogger, nestLv As Integer) Implements IExpression.Logging
            logger?.LogTrace("{}", New String(" "c, nestLv * 2) & Me.ToString())
            Me.mIndexToken.Logging(logger, nestLv + 1)
        End Sub

        ''' <summary>文字列条件を取得します。</summary>
        ''' <returns>文字列表現。</returns>
        Public Overrides Function ToString() As String
            Return $"expr:ident {If(Me.mToken?.Contents, "")} index({Me.mIndexToken})"
        End Function

    End Class

End Namespace