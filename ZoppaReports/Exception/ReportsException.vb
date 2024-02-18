Option Strict On
Option Explicit On

Namespace Exceptions

    ''' <summary>レポート例外。</summary>
    Public Class ReportsException
        Inherits Exception

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="message">例外メッセージ。</param>
        Public Sub New(message As String)
            MyBase.New(message)
        End Sub

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="message">例外メッセージ。</param>
        ''' <param name="innerEx">内部例外。</param>
        Public Sub New(message As String, innerEx As Exception)
            MyBase.New(message, innerEx)
        End Sub

    End Class

End Namespace

