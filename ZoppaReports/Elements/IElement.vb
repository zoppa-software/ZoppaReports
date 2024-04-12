Option Strict On
Option Explicit On

''' <summary>帳票要素。</summary>
Public Interface IElement

    ''' <summary>リソースプロパティを設定します。</summary>
    ''' <param name="name">リソース名。</param>
    ''' <param name="value">プロパティ値。</param>
    Sub SetProperty(name As String, value As Object)

End Interface
