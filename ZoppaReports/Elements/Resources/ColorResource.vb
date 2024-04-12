Option Strict On
Option Explicit On

Imports System.Drawing
Imports ZoppaReports.Exceptions

Namespace Resources

    ''' <summary>色リソース。</summary>
    Public NotInheritable Class ColorResource
        Implements IResources

        ' リソース名
        Private mName As String

        ' 色
        Private mColor As Color = Color.Black

#Region "properties"

        ''' <summary>リソース名を取得します。</summary>
        ''' <returns>リソース名、</returns>
        Public ReadOnly Property Name As String Implements IResources.Name
            Get
                Return Me.mName
            End Get
        End Property

        ''' <summary>リソースを取得します。</summary>
        ''' <returns>リソース。</returns>
        Private Function Contents(Of T As {Class})() As T Implements IResources.Contents
            Dim res = TryCast(Me.mColor, T)
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

        ''' <summary>リソースプロパティを設定します。</summary>
        ''' <param name="name">リソース名。</param>
        ''' <param name="value">プロパティ値。</param>
        Public Sub SetProperty(name As String, value As Object) Implements IResources.SetProperty
            Select Case name.ToLower()
                Case NameOf(Me.Color).ToLower(), "value"
                    Me.mColor = ConvertColor(value)
            End Select
        End Sub

    End Class

End Namespace
