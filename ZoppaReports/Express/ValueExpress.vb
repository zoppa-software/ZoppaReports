﻿Option Strict On
Option Explicit On

Imports Microsoft.Extensions.Logging
Imports ZoppaReports.Exceptions
Imports ZoppaReports.Tokens

Namespace Express

    ''' <summary>値式。</summary>
    Public NotInheritable Class ValueExpress
        Implements IExpression

        ' 対象トークン
        Private ReadOnly mToken As IToken

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="token">対象トークン。</param>
        Public Sub New(token As IToken)
            If token IsNot Nothing Then
                Me.mToken = token
            Else
                Throw New ReportsAnalysisException("値式の生成にNullは使用できません")
            End If
        End Sub

        ''' <summary>式を実行する。</summary>
        ''' <param name="env">環境値情報。</param>
        ''' <returns>実行結果。</returns>
        Public Function Executes(env As Environments) As IToken Implements IExpression.Executes
            Select Case Me.mToken.TokenType
                Case GetType(IdentToken)
                    Dim obj = env.GetValue(If(Me.mToken.Contents?.ToString(), ""))
                    Return ConvertValueToken(obj)

                Case Else
                    Return Me.mToken
            End Select
        End Function

        ''' <summary>ログ出力を行います。</summary>
        ''' <param name="logger">ロガー。</param>
        ''' <param name="nestLv">ネストレベル。</param>
        Public Sub Logging(logger As ILogger, nestLv As Integer) Implements IExpression.Logging
            logger?.LogTrace("{}", New String(" "c, nestLv * 2) & Me.ToString())
        End Sub

        ''' <summary>文字列条件を取得します。</summary>
        ''' <returns>文字列表現。</returns>
        Public Overrides Function ToString() As String
            Select Case Me.mToken.TokenType
                Case GetType(IdentToken)
                    Return $"expr:ident {If(Me.mToken?.Contents, "")}"
                Case Else
                    Return $"expr:{If(Me.mToken?.Contents, "null")}"
            End Select
        End Function

        ''' <summary>値をトークンに変換します。</summary>
        ''' <param name="obj">値。</param>
        ''' <returns>トークン。</returns>
        Friend Shared Function ConvertValueToken(obj As Object) As IToken
            If TypeOf obj Is String Then
                Return New StringToken(obj.ToString(), Nothing)
            ElseIf TypeOf obj Is Int32 OrElse
                   TypeOf obj Is Double OrElse
                   TypeOf obj Is Decimal OrElse
                   TypeOf obj Is UInt64 OrElse
                   TypeOf obj Is UInt32 OrElse
                   TypeOf obj Is UInt16 OrElse
                   TypeOf obj Is Single OrElse
                   TypeOf obj Is SByte OrElse
                   TypeOf obj Is Byte OrElse
                   TypeOf obj Is Int64 OrElse
                   TypeOf obj Is Int16 Then
                Return NumberToken.ConvertToken(obj)
            Else
                Return New ObjectToken(obj)
            End If
        End Function

    End Class

End Namespace