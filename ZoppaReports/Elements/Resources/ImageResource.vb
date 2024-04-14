Option Strict On
Option Explicit On

Imports System.Drawing
Imports System.IO
Imports ZoppaReports.Exceptions

Namespace Resources

    ''' <summary>画像リソース。</summary>
    Public NotInheritable Class ImageResource
        Implements IReportsResources

        ' リソース名
        Private ReadOnly mName As String

        ' 画像
        Private mImage As Image

        ' 画像ファイルパス
        Private mPath As String

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
            If Me.mImage Is Nothing Then
                Try
                    Me.mImage = Image.FromFile(Me.mPath)
                Catch ex As Exception
                    Throw New FileNotFoundException("画像ファイルの読み込みに失敗しました")
                End Try
            End If

            Dim res = TryCast(Me.mImage, T)
            If res IsNot Nothing Then
                Return res
            Else
                Throw New InvalidCastException($"リソースは{GetType(T).Name}ではありません")
            End If
        End Function

        ''' <summary>画像ファイルパスを取得します。</summary>
        ''' <returns>画像ファイルパス。</returns>
        Public ReadOnly Property Path As String
            Get
                Return Me.mPath
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
            Me.mImage?.Dispose()
            Me.mImage = Nothing
        End Sub

        ''' <summary>プロパティを設定します。</summary>
        ''' <param name="name">プロパティ名。</param>
        ''' <param name="value">プロパティ値。</param>
        ''' <returns>追加できたら真。</returns>
        Public Function SetProperty(name As String, value As Object) As Boolean Implements IReportsResources.SetProperty
            Me.mImage?.Dispose()
            Me.mImage = Nothing

            Select Case name.ToLower()
                Case NameOf(Me.Path).ToLower()
                    Me.mPath = value.ToString()
                    Return True

                Case Else
                    Return False
            End Select
        End Function

    End Class

End Namespace