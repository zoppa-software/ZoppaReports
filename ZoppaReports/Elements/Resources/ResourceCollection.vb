Option Strict On
Option Explicit On

Namespace Resources

    ''' <summary>リソースコレクション。</summary>
    Public NotInheritable Class ResourceCollection

        ' リソースリスト
        Private ReadOnly mResources As New SortedList(Of String, IReportsResources)

        ' スタイルリソースリスト
        Private ReadOnly mStyleResources As New Dictionary(Of Type, StyleResource)

        ''' <summary>リリースリストを取得します。</summary>
        ''' <returns>リリースリスト。</returns>
        ReadOnly Property Resources As IList(Of IReportsResources)
            Get
                Dim res As New List(Of IReportsResources)()
                res.AddRange(Me.mResources.Values)
                res.AddRange(Me.mStyleResources.Values)
                Return res
            End Get
        End Property

        ''' <summary>リソースをクリアします。</summary>
        Public Sub Clear()
            Me.mResources.Clear()
            Me.mStyleResources.Clear()
        End Sub

        ''' <summary>リソースを追加します。</summary>
        ''' <param name="value">リソース。</param>
        ''' <returns>リソース数。</returns>
        Public Function Add(value As IReportsResources) As Integer
            If value IsNot Nothing Then
                ' 名前付きリソースを追加
                If value.Name?.Trim() <> "" Then
                    Me.mResources.Add(value.Name.Trim(), value)
                End If

                ' スタイルリソースを追加
                Dim bstyle = TryCast(value, StyleResource)
                If bstyle IsNot Nothing AndAlso bstyle.TypeBase IsNot Nothing Then
                    Me.mStyleResources.Add(bstyle.TypeBase, bstyle)
                End If

                Return Me.mResources.Count
            Else
                Return -1
            End If
        End Function

        ''' <summary>リソースを取得します。</summary>
        ''' <typeparam name="T">リソース型。</typeparam>
        ''' <param name="name">リソース名。</param>
        ''' <returns>リソース。</returns>
        Public Function GetResource(Of T As {IReportsResources})(name As String) As T
            If Me.mResources.ContainsKey(name) Then
                Return CType(Me.mResources(name), T)
            Else
                Return Nothing
            End If
        End Function

        ''' <summary>基本スタイルを取得します。</summary>
        ''' <param name="typeBase">スタイル名。</param>
        ''' <returns>スタイルリソース。</returns>
        Public Function GetBaseStyle(typeBase As Type) As StyleResource
            If Me.mStyleResources.ContainsKey(typeBase) Then
                Return Me.mStyleResources(typeBase)
            Else
                Return Nothing
            End If
        End Function

    End Class

End Namespace