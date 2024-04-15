Option Strict On
Option Explicit On

Imports System.Drawing
Imports Microsoft.Extensions.Logging
Imports ZoppaReports.Resources
Imports ZoppaReports.Smalls

Namespace Container

    ''' <summary>キャンバス要素。</summary>
    Public NotInheritable Class CanvasElement
        Inherits CommonElement
        Implements IHasProperties

        ' 親要素
        Private ReadOnly mParent As IReportsElement

        Private Shared ReadOnly mXs As New Dictionary(Of IReportsElement, ReportsScalar)

        Private Shared ReadOnly mYs As New Dictionary(Of IReportsElement, ReportsScalar)

        Public Shared Sub SetSharedProperty(target As IReportsElement, name As String, value As Object)
            Select Case name.Trim().ToLower()
                Case "x"
                    mXs.Add(target, ConvertScalar(value))
                Case "y"
                    mYs.Add(target, ConvertScalar(value))
            End Select
        End Sub

#Region "properties"

        ''' <summary>親要素を取得します。</summary>
        Public Overrides ReadOnly Property Parent As IReportsElement
            Get
                Return Me.mParent
            End Get
        End Property

#End Region

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="parent">親要素。</param>
        Public Sub New(parent As IReportsElement)
            Me.mParent = parent
        End Sub

        ''' <summary>印刷を実行します。</summary>
        ''' <param name="graphics">グラフィックオブジェクト。</param>
        ''' <param name="clipBounds">クリップ領域。</param>
        Public Overrides Sub Draw(graphics As Graphics, clipBounds As Rectangle)
            Dim lf = CSng(Me.Margin.Left.Pixel + Me.Padding.Left.Pixel)
            Dim tp = CSng(Me.Margin.Top.Pixel + Me.Padding.Top.Pixel)

            ' マージン、パディングを適用
            graphics.TranslateTransform(lf, tp)

            ' クリップ領域を設定
            Dim wid = CInt(clipBounds.Width - Me.Margin.Right.Pixel - Me.Padding.Right.Pixel)
            Dim hig = CInt(clipBounds.Height - Me.Margin.Bottom.Pixel - Me.Padding.Bottom.Pixel)
            clipBounds = New Rectangle(0, 0, wid, hig)

            ' 子要素を描画
            Try
                Dim pervClip = graphics.ClipBounds
                For Each child In Me.mChildren
                    Dim cliped = child.ClipToBounds
                    If cliped Then
                        graphics.SetClip(clipBounds)
                    End If

                    Dim x = CInt(If(mXs.ContainsKey(child), mXs(child), ReportsScalar.Zero).Pixel)
                    Dim y = CInt(If(mYs.ContainsKey(child), mYs(child), ReportsScalar.Zero).Pixel)
                    graphics.TranslateTransform(x, y)
                    child.Draw(graphics, clipBounds)
                    graphics.TranslateTransform(-x, -y)

                    If cliped Then
                        graphics.SetClip(pervClip)
                    End If
                Next
            Catch ex As Exception
                Logger.Value?.LogError("描画に失敗しました:{Message}", ex.Message)

            Finally
                graphics.TranslateTransform(-lf, -tp)
            End Try
        End Sub

        ''' <summary>プロパティを設定します。</summary>
        ''' <param name="name">プロパティ名。</param>
        ''' <param name="value">プロパティ値。</param>
        ''' <returns>追加できたら真。</returns>
        Public Function SetProperty(name As String, value As Object) As Boolean Implements IHasProperties.SetProperty
            Select Case name.ToLower()
                Case Else
                    If Me.SetCommonProperty(name, value) Then
                        Return True
                    End If
                    Return False
            End Select
        End Function

    End Class

End Namespace
