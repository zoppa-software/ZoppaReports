Option Strict On
Option Explicit On

''' <summary>設定プロパティを持つ要素インターフェース。</summary>
Public Interface IHasProperties

    ''' <summary>プロパティを設定します。</summary>
    ''' <param name="name">プロパティ名。</param>
    ''' <param name="value">プロパティ値。</param>
    ''' <returns>追加できたら真。</returns>
    Function SetProperty(name As String, value As Object) As Boolean

End Interface
