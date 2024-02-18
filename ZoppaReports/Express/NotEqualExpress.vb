Option Strict On
Option Explicit On

Imports Microsoft.Extensions.Logging
Imports ZoppaReports.Exceptions
Imports ZoppaReports.Tokens

Namespace Express

    ''' <summary>不等号式。</summary>
    Public NotInheritable Class NotEqualExpress
        Implements IExpression

        ' 左辺式
        Private ReadOnly mTml As IExpression

        ' 右辺式
        Private ReadOnly mTmr As IExpression

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="tml">左辺式。</param>
        ''' <param name="tmr">右辺式。</param>
        Public Sub New(tml As IExpression, tmr As IExpression)
            If tml IsNot Nothing AndAlso tmr IsNot Nothing Then
                Me.mTml = tml
                Me.mTmr = tmr
            Else
                Throw New ReportsAnalysisException("不等号式の生成にNullは使用できません")
            End If
        End Sub

        ''' <summary>式を実行する。</summary>
        ''' <param name="env">環境値情報。</param>
        ''' <returns>実行結果。</returns>
        Public Function Executes(env As Environments) As IToken Implements IExpression.Executes
            Dim tml = Me.mTml.Executes(env)
            Dim tmr = Me.mTmr.Executes(env)

            Dim nml = TryCast(tml, NumberToken)
            Dim nmr = TryCast(tmr, NumberToken)
            Dim bval As Boolean
            If nml IsNot Nothing AndAlso nmr IsNot Nothing Then
                ' 数値比較
                bval = nml.EqualCondition(nmr)
            ElseIf tml?.Contents Is Nothing AndAlso tmr.Contents Is Nothing Then
                ' 両方 Null比較
                bval = True
            Else
                bval = If(tml.Contents?.Equals(tmr.Contents), False)
            End If
            Return If(bval, CType(FalseToken.Value, IToken), TrueToken.Value)
        End Function

        ''' <summary>ログ出力を行います。</summary>
        ''' <param name="logger">ロガー。</param>
        ''' <param name="nestLv">ネストレベル。</param>
        Public Sub Logging(logger As ILogger, nestLv As Integer) Implements IExpression.Logging
            logger?.LogTrace(New String(" "c, nestLv * 2) & Me.ToString())
            Me.mTml.Logging(logger, nestLv + 1)
            Me.mTmr.Logging(logger, nestLv + 1)
        End Sub

        ''' <summary>文字列条件を取得します。</summary>
        ''' <returns>文字列表現。</returns>
        Public Overrides Function ToString() As String
            Return "expr:<>"
        End Function

    End Class

End Namespace