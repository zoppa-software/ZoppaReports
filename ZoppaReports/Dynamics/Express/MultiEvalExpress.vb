﻿Option Strict On
Option Explicit On

Imports Microsoft.Extensions.Logging
Imports ZoppaReports.Exceptions
Imports ZoppaReports.Tokens

Namespace Express

    ''' <summary>三項演算子式。</summary>
    Public NotInheritable Class MultiEvalExpress
        Implements IExpression

        ' 条件式
        Private ReadOnly mCondition As IExpression

        ' 真式
        Private ReadOnly mTrue As IExpression

        ' 偽式
        Private ReadOnly mFalse As IExpression

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="condition">条件式。</param>
        ''' <param name="trueTerm">真式。</param>
        ''' <param name="falseTerm">偽式。</param>
        Public Sub New(condition As IExpression, trueTerm As IExpression, falseTerm As IExpression)
            If condition IsNot Nothing AndAlso trueTerm IsNot Nothing AndAlso falseTerm IsNot Nothing Then
                Me.mCondition = condition
                Me.mTrue = trueTerm
                Me.mFalse = falseTerm
            Else
                Throw New ReportsAnalysisException("三項演算子式の生成にNullは使用できません")
            End If
        End Sub

        ''' <summary>式を実行する。</summary>
        ''' <param name="env">環境値情報。</param>
        ''' <returns>実行結果。</returns>
        Public Function Executes(env As Environments) As IToken Implements IExpression.Executes
            Dim cond = Me.mCondition.Executes(env)
            If TypeOf cond.Contents Is Boolean Then
                If Convert.ToBoolean(cond.Contents) Then
                    Return Me.mTrue.Executes(env)
                Else
                    Return Me.mFalse.Executes(env)
                End If
            Else
                Throw New InvalidOperationException($"三項演算の第一式は真偽値の結果が必要です")
            End If
        End Function

        ''' <summary>ログ出力を行います。</summary>
        ''' <param name="logger">ロガー。</param>
        ''' <param name="nestLv">ネストレベル。</param>
        Public Sub Logging(logger As ILogger, nestLv As Integer) Implements IExpression.Logging
            logger?.LogTrace("{}", New String(" "c, nestLv * 2) & Me.ToString())
            Me.mCondition.Logging(logger, nestLv + 1)
            Me.mTrue.Logging(logger, nestLv + 1)
            Me.mFalse.Logging(logger, nestLv + 1)
        End Sub

        ''' <summary>文字列条件を取得します。</summary>
        ''' <returns>文字列表現。</returns>
        Public Overrides Function ToString() As String
            Return "expr:?"
        End Function

    End Class

End Namespace