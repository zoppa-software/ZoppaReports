Option Strict On
Option Explicit On

Namespace Resources

    ''' <summary>リソースインターフェース。</summary>
    Public Interface IResources
        Inherits IElement, IDisposable

        ''' <summary>リソース名を取得します。</summary>
        ''' <returns>リソース名、</returns>
        ReadOnly Property Name As String

        ''' <summary>リソースを取得します。</summary>
        ''' <returns>リソース。</returns>
        Function Contents(Of T As {Class})() As T

    End Interface

End Namespace
