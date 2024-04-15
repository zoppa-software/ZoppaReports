Option Strict On
Option Explicit On

Imports System.Drawing
Imports ZoppaReports.Resources
Imports ZoppaReports.Smalls

Namespace Container

    ''' <summary>共通要素。</summary>
    Public MustInherit Class CommonElement
        Implements IReportsElement

        ' リソースコレクション
        Protected ReadOnly mResCollection As New ResourceCollection

        ' 子要素コレクション
        Protected ReadOnly mChildren As New List(Of IReportsElement)

        ' マージン
        Private mMargin As New ReportsMargin()

        ' パディング
        Private mPadding As New ReportsPadding()

        ' クリップする、しない
        Private mClipToBounds As Boolean = True

        ''' <summary>親要素を取得します。</summary>
        Public MustOverride ReadOnly Property Parent As IReportsElement Implements IReportsElement.Parent

        ''' <summary>子要素リストを取得します。</summary>
        Public ReadOnly Property Children As IList(Of IReportsElement) Implements IReportsElement.Children
            Get
                Return Me.mChildren
            End Get
        End Property

        ''' <summary>リソースリストを取得します。</summary>
        Public ReadOnly Property Resources As IList(Of IReportsResources) Implements IReportsElement.Resources
            Get
                Return Me.mResCollection.Resources
            End Get
        End Property

        ''' <summary>マージンを取得します。</summary>
        Public ReadOnly Property Margin As ReportsMargin Implements IReportsElement.Margin
            Get
                Return Me.mMargin
            End Get
        End Property

        ''' <summary>パディングを取得します。</summary>
        Public ReadOnly Property Padding As ReportsPadding Implements IReportsElement.Padding
            Get
                Return Me.mPadding
            End Get
        End Property

        ''' <summary>クリップする、しないを取得します。</summary>
        Public ReadOnly Property ClipToBounds As Boolean Implements IReportsElement.ClipToBounds
            Get
                Return Me.mClipToBounds
            End Get
        End Property

        ''' <summary>子要素を追加します。</summary>
        ''' <param name="child">子要素。</param>
        Public Sub AddChildElement(child As IReportsElement) Implements IReportsElement.AddChildElement
            Me.mChildren.Add(child)
        End Sub

        ''' <summary>リソースを追加します。</summary>
        ''' <param name="addRes">リソース。</param>
        Public Sub AddResources(addRes As IReportsResources) Implements IReportsElement.AddResources
            Me.mResCollection.Add(addRes)
        End Sub

        ''' <summary>リソースを取得します。</summary>
        ''' <param name="name">リソース名。</param>
        ''' <returns>リソース。</returns>
        Public Function GetResource(Of T As {IReportsResources})(name As String) As T Implements IReportsElement.GetResource
            Return Me.mResCollection.GetResource(Of T)(name)
        End Function

        ''' <summary>基本スタイルを取得します。</summary>
        ''' <param name="typeBase">適用する型。</param>
        ''' <returns>スタイルリソース。</returns>
        Public Function GetBaseStyle(typeBase As Type) As StyleResource Implements IReportsElement.GetBaseStyle
            Return Me.mResCollection.GetBaseStyle(typeBase)
        End Function

        ''' <summary>リソースを解放します。</summary>
        Public Overridable Sub Dispose() Implements IDisposable.Dispose
            Me.mResCollection.Dispose()
            For Each child As IReportsElement In Me.mChildren
                child.Dispose()
            Next
        End Sub

        ''' <summary>印刷を実行します。</summary>
        ''' <param name="graphics">グラフィックオブジェクト。</param>
        ''' <param name="clipBounds">クリップ領域。</param>
        Public MustOverride Sub Draw(graphics As Graphics, clipBounds As Rectangle) Implements IReportsElement.Draw

        ''' <summary>プロパティを設定します。</summary>
        ''' <param name="name">プロパティ名。</param>
        ''' <param name="value">プロパティ値。</param>
        ''' <returns>追加できたら真。</returns>
        Protected Function SetCommonProperty(name As String, value As Object) As Boolean
            Select Case name.ToLower()
                Case NameOf(Me.Margin).ToLower()
                    Me.mMargin = CType(value.ToString(), ReportsMargin)
                    Return True

                Case NameOf(Me.Padding).ToLower()
                    Me.mPadding = CType(value, ReportsPadding)
                    Return True

                Case NameOf(Me.ClipToBounds).ToLower()
                    Me.mClipToBounds = Convert.ToBoolean(value)
                    Return True

                Case Else
                    Return False
            End Select
        End Function

    End Class

End Namespace

