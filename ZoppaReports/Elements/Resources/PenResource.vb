Option Strict On
Option Explicit On

Imports System.Drawing
Imports ZoppaReports.Exceptions

Namespace Resources

    ''' <summary>ペンリソース。</summary>
    Public NotInheritable Class PenResource
        Implements IReportsResources

        ' リソース名
        Private ReadOnly mName As String

        ' ブラシ
        Private mPen As Pen

        ' 色
        Private mColor As Color = Color.Black

        ' 太さ
        Private mWidth As Single = 1.0F

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
            If Me.mPen Is Nothing Then
                Me.mPen = New Pen(Me.mColor, Me.mWidth)
            End If

            Dim res = TryCast(Me.mPen, T)
            If res IsNot Nothing Then
                Return res
            Else
                Throw New InvalidCastException($"リソースは{GetType(T).Name}ではありません")
            End If
        End Function

        ''' <summary>色を取得します。</summary>
        ''' <returns>色。</returns>
        Public ReadOnly Property Color As Color
            Get
                Return Me.mColor
            End Get
        End Property

        ''' <summary>太さを取得します。</summary>
        ''' <returns>太さ。</returns>
        Public ReadOnly Property Width As Single
            Get
                Return Me.mWidth
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
            Me.mPen?.Dispose()
            Me.mPen = Nothing
        End Sub

        ''' <summary>プロパティを設定します。</summary>
        ''' <param name="name">プロパティ名。</param>
        ''' <param name="value">プロパティ値。</param>
        ''' <returns>追加できたら真。</returns>
        Public Function SetProperty(name As String, value As Object) As Boolean Implements IReportsResources.SetProperty
            Me.mPen?.Dispose()
            Me.mPen = Nothing

            Select Case name.ToLower()
                Case NameOf(Me.Color).ToLower()
                    Me.mColor = ConvertColor(value)
                    Return True

                Case NameOf(Me.Width).ToLower()
                    Try
                        Me.mWidth = Convert.ToSingle(value)
                        Return True
                    Catch ex As Exception
                        Throw New ReportsAnalysisException($"{NameOf(Me.Width)}プロパティに{value.GetType().Name}を設定できません", ex)
                    End Try

                Case Else
                    Return False
            End Select
        End Function

    End Class

End Namespace

