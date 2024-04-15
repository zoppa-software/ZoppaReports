Option Strict On
Option Explicit On

Imports System.Drawing
Imports ZoppaReports.Container
Imports ZoppaReports.Resources
Imports ZoppaReports.Smalls

''' <summary>レポート要素。</summary>
Public NotInheritable Class ReportsElement
    Inherits CommonElement
    Implements IHasProperties

    ' ページサイズ
    Private mSize As ReportsSize = ReportsSize.A4

    ' ページ向き
    Private mOrientation As ReportsOrientation = ReportsOrientation.Portrait

    ' 幅
    Private mWidth As ReportsScalar = ReportsScalar.Zero

    ' 高さ
    Private mHeight As ReportsScalar = ReportsScalar.Zero

#Region "properties"

    ''' <summary>親要素を取得します。</summary>
    Public Overrides ReadOnly Property Parent As IReportsElement
        Get
            Return Nothing
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

    ''' <summary>印刷を実行します。</summary>
    ''' <param name="graphics">グラフィックオブジェクト。</param>
    ''' <param name="clipBounds">クリップ領域。</param>
    Public Overrides Sub Draw(graphics As Graphics, clipBounds As Rectangle)
        ' マージン、パディングを適用
        graphics.TranslateTransform(
            CSng(Me.Margin.Left.Pixel + Me.Padding.Left.Pixel),
            CSng(Me.Margin.Top.Pixel + Me.Padding.Top.Pixel)
        )

        ' クリップ領域を設定
        Dim wid = CInt(Me.mSize.Width.Pixel)
        Dim hig = CInt(Me.mSize.Height.Pixel)
        clipBounds = If(
            Me.mOrientation = ReportsOrientation.Landscape,
            New Rectangle(0, 0, hig, wid),
            New Rectangle(0, 0, wid, hig)
        )

        ' 子要素を描画
        For Each child In Me.mChildren
            Dim cliped = child.ClipToBounds
            If cliped Then
                graphics.SetClip(clipBounds)
            End If

            child.Draw(graphics, clipBounds)

            If cliped Then
                graphics.ResetClip()
            End If
        Next
    End Sub

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
                If Me.SetCommonProperty(name, value) Then
                    Return True
                End If
                Return False
        End Select
    End Function

End Class
