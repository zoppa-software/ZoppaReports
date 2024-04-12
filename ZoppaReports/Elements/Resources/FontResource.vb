Option Strict On
Option Explicit On

Imports System.Drawing
Imports ZoppaReports.Exceptions

Namespace Resources

    ''' <summary>フォントリソース。</summary>
    Public NotInheritable Class FontResource
        Implements IResources

        ' リソース名
        Private ReadOnly mName As String

        ' フォント
        Private mFont As Font

        ' フォントファミリ名
        Private mFamily As String = "Meiryo UI"

        ' フォントサイズ
        Private mEmSize As Single = 12.0F

        ' フォントスタイル
        Private mStyle As FontStyle = FontStyle.Regular

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
            If Me.mFont Is Nothing Then
                Me.mFont = New Font(Me.Family, Me.EmSize, Me.Style)
            End If

            Dim res = TryCast(Me.mFont, T)
            If res IsNot Nothing Then
                Return res
            Else
                Throw New InvalidCastException($"リソースは{GetType(T).Name}ではありません")
            End If
        End Function

        ''' <summary>フォントファミリ名を取得、設定します。</summary>
        Public ReadOnly Property Family As String
            Get
                Return Me.mFamily
            End Get
        End Property

        ''' <summary>フォントサイズを取得、設定します。</summary>
        Public ReadOnly Property EmSize As Single
            Get
                Return Me.mEmSize
            End Get
        End Property

        ''' <summary>フォントスタイルを取得、設定します。</summary>
        Public ReadOnly Property Style As FontStyle
            Get
                Return Me.mStyle
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
            Me.mFont?.Dispose()
            Me.mFont = Nothing
        End Sub

        ''' <summary>リソースプロパティを設定します。</summary>
        ''' <param name="name">リソース名。</param>
        ''' <param name="value">プロパティ値。</param>
        Public Sub SetProperty(name As String, value As Object) Implements IResources.SetProperty
            Me.mFont?.Dispose()
            Me.mFont = Nothing

            Select Case name.ToLower()
                Case NameOf(Me.Family).ToLower()
                    Me.mFamily = value.ToString()

                Case NameOf(Me.EmSize).ToLower()
                    Try
                        Me.mEmSize = Convert.ToSingle(value)
                    Catch ex As Exception
                        Throw New ReportsAnalysisException($"{NameOf(Me.EmSize)}プロパティに{value.GetType().Name}を設定できません", ex)
                    End Try

                Case NameOf(Me.Style).ToLower()
                    If Not [Enum].TryParse(Of FontStyle)(value.ToString(), Me.mStyle) Then
                        Throw New ReportsAnalysisException($"{NameOf(Me.EmSize)}プロパティに{value.GetType().Name}を設定できません")
                    End If
            End Select
        End Sub

    End Class

End Namespace