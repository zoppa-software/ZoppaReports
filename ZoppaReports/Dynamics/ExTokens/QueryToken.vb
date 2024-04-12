﻿Option Strict On
Option Explicit On

Imports ZoppaReports.Tokens

Namespace ExTokens

    ''' <summary>クエリトークン。</summary>
    Public NotInheritable Class QueryToken
        Implements IToken

        ' 出力する文字列
        Private ReadOnly mValue As String

        ''' <summary>格納されている値を取得する。</summary>
        ''' <returns>格納値。</returns>
        Public ReadOnly Property Contents As Object Implements IToken.Contents
            Get
                Return Me.mValue
            End Get
        End Property

        ''' <summary>トークン型を取得する。</summary>
        ''' <returns>トークン型。</returns>
        Public ReadOnly Property TokenType As Type Implements IToken.TokenType
            Get
                Return GetType(QueryToken)
            End Get
        End Property

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="value">対象文字列。</param>
        Public Sub New(value As String)
            Me.mValue = value
        End Sub

        ''' <summary>文字列条件を取得します。</summary>
        ''' <returns>文字列表現。</returns>
        Public Overrides Function ToString() As String
            Return Me.mValue
        End Function

    End Class

End Namespace


