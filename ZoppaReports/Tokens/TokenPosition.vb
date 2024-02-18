﻿Option Strict On
Option Explicit On

Imports ZoppaReports.Tokens

Namespace Tokens

    ''' <summary>トークン位置情報です。</summary>
    Public Structure TokenPosition

        ' 参照しているトークン
        Private ReadOnly mToken As IToken

        ''' <summary>トークン位置。</summary>
        Public ReadOnly Position As Integer

        '''' <summary>格納されている値を取得する。</summary>
        '''' <returns>格納値。</returns>
        'Public ReadOnly Property Contents As Object
        '    Get
        '        Return Me.mToken.Contents
        '    End Get
        'End Property

        ''' <summary>トークン型を取得する。</summary>
        ''' <returns>トークン型。</returns>
        Public ReadOnly Property TokenType As Type
            Get
                Return If(Me.mToken?.TokenType, Nothing)
            End Get
        End Property

        '''' <summary>コントロールトークンならば真を返します。</summary>
        '''' <returns>コントロールトークンならば真。</returns>
        'Public ReadOnly Property IsControlToken As Boolean
        '    Get
        '        Return TypeOf Me.mToken Is IControlToken
        '    End Get
        'End Property

        '''' <summary>トークンが空白文字ならば真を返します。</summary>
        '''' <returns>トークンが空白文字ならば真。</returns>
        'Public ReadOnly Property IsWhiteSpace As Boolean
        '    Get
        '        Return Me.mToken.IsWhiteSpace
        '    End Get
        'End Property

        '''' <summary>トークンが改行文字ならば真を返します。</summary>
        '''' <returns>トークンが改行文字ならば真。</returns>
        'Public ReadOnly Property IsCrLf As Boolean
        '    Get
        '        Return Me.mToken.IsCrLf
        '    End Get
        'End Property

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="src">参照するトークン。</param>
        ''' <param name="pos">トークン位置。</param>
        Public Sub New(src As IToken, pos As Integer)
            Me.mToken = src
            Me.Position = pos
        End Sub

        ''' <summary>型を指定してトークンを取得します。</summary>
        ''' <typeparam name="T">取得する型。</typeparam>
        ''' <returns>トークン情報。</returns>
        Friend Function GetToken(Of T As {Class, IToken})() As T
            Return TryCast(Me.mToken, T)
        End Function

        ''' <summary>文字列表現を取得する。</summary>
        ''' <returns>文字列。</returns>
        Public Overrides Function ToString() As String
            Return Me.mToken.ToString()
        End Function

    End Structure

End Namespace
