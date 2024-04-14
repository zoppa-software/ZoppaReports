Option Strict On
Option Explicit On

Imports System.Drawing
Imports ZoppaReports.Resources

Namespace Draw

    ''' <summary>ラベル要素。</summary>
    Public NotInheritable Class LabelElement
        Implements IReportsElement, IHasProperties

        ' デフォルトフォント
        Private Shared mDefaultFont As Font = New Font("Meiryo", 12)

        ' 親要素
        Private ReadOnly mParent As IReportsElement

        ' リソースコレクション
        Private ReadOnly mResCollection As New ResourceCollection

        ' 子要素コレクション
        Private ReadOnly mChildren As New List(Of IReportsElement)

        ' テキスト
        Private mText As String = ""

        ' フォント
        Private mFont As Font = mDefaultFont

        ' ブラシ
        Private mBrush As Brush = Brushes.Black

#Region "properties"

        ''' <summary>親要素を取得します。</summary>
        Public ReadOnly Property Parent As IReportsElement Implements IReportsElement.Parent
            Get
                Return Me.mParent
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

        ''' <summary>テキストを取得します。</summary>
        Public ReadOnly Property Text As String
            Get
                Return Me.mText
            End Get
        End Property

        Public ReadOnly Property Font As Font
            Get
                Return Me.mFont
            End Get
        End Property

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
            graphics.DrawString(Me.mText, Me.mFont, Brushes.Black, 0, 0)

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
                    Return False
            End Select
        End Function

    End Class

End Namespace