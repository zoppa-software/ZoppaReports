Option Strict On
Option Explicit On

Imports Microsoft.Extensions.Logging
Imports ZoppaReports.Exceptions
Imports ZoppaReports.Tokens

Namespace Express

    ''' <summary>前置き式。</summary>
    Public NotInheritable Class UnaryExpress
        Implements IExpression

        ' 対象トークン
        Private ReadOnly mToken As IToken

        ' 対象式
        Private ReadOnly mValue As IExpression

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="token">対象トークン。</param>
        Public Sub New(token As IToken, factor As IExpression)
            If token IsNot Nothing AndAlso factor IsNot Nothing Then
                Me.mToken = token
                Me.mValue = factor
            Else
                Throw New ReportsAnalysisException("前置き式の生成にNullは使用できません")
            End If
        End Sub

        ''' <summary>式を実行する。</summary>
        ''' <param name="env">環境値情報。</param>
        ''' <returns>実行結果。</returns>
        Public Function Executes(env As Environments) As IToken Implements IExpression.Executes
            Dim tkn = Me.mValue.Executes(env)

            Select Case Me.mToken.TokenType
                Case GetType(PlusToken)
                    If tkn?.TokenType = GetType(NumberToken) Then
                        Return tkn
                    Else
                        Throw New ReportsAnalysisException("前置き+が数字の前に置かれていません")
                    End If

                Case GetType(MinusToken)
                    If tkn?.TokenType = GetType(NumberToken) Then
                        Return CType(tkn, NumberToken).SignChange()
                    Else
                        Throw New ReportsAnalysisException("前置き-が数字の前に置かれていません")
                    End If

                Case GetType(NotToken)
                    Dim bval = Convert.ToBoolean(tkn?.Contents)
                    Return If(bval, CType(FalseToken.Value, IToken), TrueToken.Value)

                Case Else
                    Throw New ReportsAnalysisException("有効な前置き式ではありません")
            End Select
        End Function

        ''' <summary>ログ出力を行います。</summary>
        ''' <param name="logger">ロガー。</param>
        ''' <param name="nestLv">ネストレベル。</param>
        Public Sub Logging(logger As ILogger, nestLv As Integer) Implements IExpression.Logging
            logger?.LogTrace(New String(" "c, nestLv * 2) & Me.ToString())
            Me.mValue.Logging(logger, nestLv + 1)
        End Sub

        ''' <summary>文字列条件を取得します。</summary>
        ''' <returns>文字列表現。</returns>
        Public Overrides Function ToString() As String
            Select Case Me.mToken.TokenType
                Case GetType(PlusToken)
                    Return $"expr:pre +"
                Case GetType(MinusToken)
                    Return $"expr:pre +"
                Case Else
                    Return $"expr:pre not"
            End Select
        End Function

    End Class

End Namespace