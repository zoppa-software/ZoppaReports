Option Strict On
Option Explicit On

Imports Microsoft.Extensions.Logging
Imports ZoppaReports.Exceptions
Imports ZoppaReports.Tokens

Namespace Express

    ''' <summary>メソッド式。</summary>
    Public NotInheritable Class MethodEvalExpress
        Implements IExpression

        ' メソッド名
        Private ReadOnly mMethod As IdentToken

        ' 真式
        Private ReadOnly mArgs As IExpression

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="name">条件式。</param>
        ''' <param name="args">真式。</param>
        Public Sub New(name As IdentToken, args As IExpression)
            If name IsNot Nothing AndAlso args IsNot Nothing Then
                Me.mMethod = name
                Me.mArgs = args
            Else
                Throw New ReportsAnalysisException("メソッド式の生成にNullは使用できません")
            End If
        End Sub

        ''' <summary>式を実行する。</summary>
        ''' <param name="env">環境値情報。</param>
        ''' <returns>実行結果。</returns>
        Public Function Executes(env As Environments) As IToken Implements IExpression.Executes
            Dim mtd = Environments.GetMethod(If(Me.mMethod.Contents?.ToString(), ""))
            Dim prm = CType(Me.mArgs.Executes(env).Contents, Object())
            Return mtd.Invoke(prm)
        End Function

        ''' <summary>ログ出力を行います。</summary>
        ''' <param name="logger">ロガー。</param>
        ''' <param name="nestLv">ネストレベル。</param>
        Public Sub Logging(logger As ILogger, nestLv As Integer) Implements IExpression.Logging
            logger?.LogTrace("{}", New String(" "c, nestLv * 2) & Me.ToString())
            Me.mArgs.Logging(logger, nestLv + 1)
        End Sub

        ''' <summary>文字列条件を取得します。</summary>
        ''' <returns>文字列表現。</returns>
        Public Overrides Function ToString() As String
            Return $"expr:method {Me.mMethod}"
        End Function

    End Class

End Namespace