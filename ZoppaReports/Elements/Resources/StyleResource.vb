Option Strict On
Option Explicit On

Imports System.Drawing
Imports ZoppaReports.Exceptions

Namespace Resources

    ''' <summary>スタイルリソース。</summary>
    Public NotInheritable Class StyleResource
        Implements IResources

        ' リソース名
        Private ReadOnly mName As String

        ' プロパティリスト
        Private ReadOnly mProperties As New SortedDictionary(Of String, Object)

#Region "properties"

        ''' <summary>リソース名を取得します。</summary>
        ''' <returns>リソース名、</returns>
        Public ReadOnly Property Name As String Implements IResources.Name
            Get
                Return Me.mName
            End Get
        End Property

        ''' <summary>リソースを取得します。</summary>
        ''' <returns>リソース。</returns>
        Private Function Contents(Of T As {Class})() As T Implements IResources.Contents
            Dim res = TryCast(Me, T)
            If res IsNot Nothing Then
                Return res
            Else
                Throw New InvalidCastException($"リソースは{GetType(T).Name}ではありません")
            End If
        End Function

#End Region

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="nm">リソース名。</param>
        Public Sub New(nm As String)
            Me.mName = nm
        End Sub

        ''' <summary>リソースを解放します。</summary>
        Public Sub Dispose() Implements IDisposable.Dispose
            ' 空実装
        End Sub

        ''' <summary>リソースプロパティを設定します。</summary>
        ''' <param name="name">リソース名。</param>
        ''' <param name="value">プロパティ値。</param>
        Public Sub SetProperty(name As String, value As Object) Implements IResources.SetProperty
            ' 空実装
        End Sub

        ''' <summary>プロパティを展開します。</summary>
        ''' <param name="target">対象要素。</param>
        Public Sub ExpandProperties(target As IElement)
            For Each prop In Me.mProperties
                target.SetProperty(prop.Key, prop.Value)
            Next
        End Sub

    End Class

End Namespace
