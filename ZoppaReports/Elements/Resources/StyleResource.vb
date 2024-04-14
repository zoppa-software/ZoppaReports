Option Strict On
Option Explicit On

Imports System.Drawing
Imports System.Reflection
Imports ZoppaReports.Exceptions

Namespace Resources

    ''' <summary>スタイルリソース。</summary>
    Public NotInheritable Class StyleResource
        Implements IReportsResources

        ' リソース名
        Private ReadOnly mName As String

        ' プロパティリスト
        Private ReadOnly mProperties As New SortedDictionary(Of String, Object)

        ' 対象型情報
        Private mBase As Type = Nothing

#Region "properties"

        ''' <summary>リソース名を取得します。</summary>
        ''' <returns>リソース名、</returns>
        Public ReadOnly Property Name As String Implements IReportsResources.Name
            Get
                Return Me.mName
            End Get
        End Property

        ''' <summary>リソースを取得します。</summary>
        ''' <returns>リソース。</returns>
        Private Function Contents(Of T As {Class})() As T Implements IReportsResources.Contents
            Dim res = TryCast(Me, T)
            If res IsNot Nothing Then
                Return res
            Else
                Throw New InvalidCastException($"リソースは{GetType(T).Name}ではありません")
            End If
        End Function

        ''' <summary>適用する要素の型を取得します。</summary>
        ''' <returns>太さ。</returns>
        Public ReadOnly Property TypeBase As Type
            Get
                Return Me.mBase
            End Get
        End Property

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

        ''' <summary>プロパティを設定します。</summary>
        ''' <param name="name">プロパティ名。</param>
        ''' <param name="value">プロパティ値。</param>
        ''' <returns>追加できたら真。</returns>
        Public Function SetProperty(name As String, value As Object) As Boolean Implements IReportsResources.SetProperty
            Select Case name.ToLower()
                Case NameOf(Me.TypeBase).ToLower(), "base"
                    Me.mBase = GetElementType(value.ToString())
                    Return True

                Case Else
                    Return False
            End Select
        End Function

        ''' <summary>プロパティを展開します。</summary>
        ''' <param name="target">対象要素。</param>
        Public Sub ExpandProperties(target As IHasProperties)
            For Each prop In Me.mProperties
                target.SetProperty(prop.Key, prop.Value)
            Next
        End Sub

        ''' <summary>プロパティを追加します。</summary>
        ''' <param name="prop">プロパティ。</param>
        Friend Sub AddSetter(prop As (key As String, val As Object))
            If prop.key <> "" Then
                Me.mProperties.Add(prop.key, prop.val)
            End If
        End Sub
    End Class

End Namespace
