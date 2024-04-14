Option Strict On
Option Explicit On

Imports System.Drawing
Imports ZoppaReports.Resources
Imports ZoppaReports.Settings

''' <summary>レポート要素。</summary>
Public NotInheritable Class ReportsElement
    Implements IReportsElement, IHasProperties

    ' リソースコレクション
    Private ReadOnly mResCollection As New ResourceCollection

    ' 子要素コレクション
    Private ReadOnly mChildren As New List(Of IReportsElement)

    ' ページサイズ
    Private mSize As ReportsSize = ReportsSize.A4

    ' ページ向き
    Private mOrientation As ReportsOrientation = ReportsOrientation.Portrait

    ' 幅
    Private mWidth As ReportsCoord = ReportsCoord.Zero

    ' 高さ
    Private mHeight As ReportsCoord = ReportsCoord.Zero

#Region "properties"

    ''' <summary>親要素を取得します。</summary>
    Public ReadOnly Property Parent As IReportsElement Implements IReportsElement.Parent
        Get
            Return Nothing
        End Get
    End Property

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

    ''' <summary>ページサイズを取得します。</summary>
    Public ReadOnly Property PaperSize As ReportsSize
        Get
            Return Me.mSize
        End Get
    End Property

    ''' <summary>ページ向きを取得します。</summary>
    Public ReadOnly Property PaperOrientation As ReportsOrientation
        Get
            Return Me.mOrientation
        End Get
    End Property

#End Region

    ''' <summary>コンストラクタ。</summary>
    Public Sub New()
        ' 空実装
    End Sub

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

    ''' <summary>印刷を実行します。</summary>
    ''' <param name="graphics">グラフィックオブジェクト。</param>
    Public Sub Draw(graphics As Graphics) Implements IReportsElement.Draw
        If Me.mOrientation = ReportsOrientation.Landscape Then
            graphics.SetClip(New Rectangle(0, 0, CInt(Me.mSize.Height.Pixel), CInt(Me.mSize.Width.Pixel)))
        Else
            graphics.SetClip(New Rectangle(0, 0, CInt(Me.mSize.Width.Pixel), CInt(Me.mSize.Height.Pixel)))
        End If

        For Each child In Me.mChildren
            child.Draw(graphics)
        Next
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

    ''' <summary>プロパティを設定します。</summary>
    ''' <param name="name">プロパティ名。</param>
    ''' <param name="value">プロパティ値。</param>
    ''' <returns>追加できたら真。</returns>
    Public Function SetProperty(name As String, value As Object) As Boolean Implements IHasProperties.SetProperty
        Select Case name.ToLower()
            Case NameOf(Me.PaperSize).ToLower(), "size"
                If value.GetType() Is GetType(ReportsSize) Then
                    Me.mSize = CType(value, ReportsSize)
                Else
                    Me.mSize = CTypeDynamic(Of ReportsSize)(value.ToString())
                End If
                Return True

            Case NameOf(Me.PaperOrientation).ToLower(), "orientation"
                If value.GetType() Is GetType(ReportsOrientation) Then
                    Me.mOrientation = CType(value, ReportsOrientation)
                Else
                    Me.mOrientation = CTypeDynamic(Of ReportsOrientation)(value.ToString())
                End If
                Return True

            Case Else
                Return False
        End Select
    End Function

End Class
