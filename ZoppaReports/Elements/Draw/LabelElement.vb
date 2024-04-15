Option Strict On
Option Explicit On

Imports System.Drawing
Imports ZoppaReports.Container
Imports ZoppaReports.Resources
Imports ZoppaReports.Smalls

Namespace Draw

    ''' <summary>ラベル要素。</summary>
    Public NotInheritable Class LabelElement
        Inherits CommonElement
        Implements IHasProperties

        ' 親要素
        Private ReadOnly mParent As IReportsElement

        ' テキスト
        Private mText As String = ""

        ' フォント
        Private mFont As Font = DefaultFont

        ' ブラシ
        Private mBrush As Brush = Brushes.Black

#Region "properties"

        ''' <summary>親要素を取得します。</summary>
        Public Overrides ReadOnly Property Parent As IReportsElement
            Get
                Return Me.mParent
            End Get
        End Property

        ''' <summary>テキストを取得します。</summary>
        Public ReadOnly Property Text As String
            Get
                Return Me.mText
            End Get
        End Property

        ''' <summary>フォントを取得します。</summary>
        Public ReadOnly Property Font As Font
            Get
                Return Me.mFont
            End Get
        End Property

        ''' <summary>ブラシを取得します。</summary>
        Public ReadOnly Property Brush As Brush
            Get
                Return Me.mBrush
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
            graphics.DrawString(Me.mText, Me.mFont, Brushes.Black, 0, 0)

            For Each child In Me.mChildren
                child.Draw(graphics, clipBounds)
            Next
        End Sub

        ''' <summary>プロパティを設定します。</summary>
        ''' <param name="name">プロパティ名。</param>
        ''' <param name="value">プロパティ値。</param>
        ''' <returns>追加できたら真。</returns>
        Public Function SetProperty(name As String, value As Object) As Boolean Implements IHasProperties.SetProperty
            Select Case name.ToLower()
                Case NameOf(Me.Text).ToLower(), "caption"
                    Me.mText = value.ToString()
                    Return True

                Case NameOf(Me.Font).ToLower()
                    Me.mFont = ConvertFont(value)
                    Return True

                Case NameOf(Me.Brush).ToLower()
                    Me.mBrush = ConvertBrush(value)
                    Return True

                Case Else
                    If Me.SetCommonProperty(name, value) Then
                        Return True
                    End If
                    If name.IndexOf("."c) > 0 Then
                        Dim names = name.Split("."c)
                        GetElementType(names(0)).GetMethod("SetSharedProperty").Invoke(Nothing, {Me, names(1), value})
                    End If
                    Return False
            End Select
        End Function

    End Class

End Namespace