Option Strict On
Option Explicit On

Imports System.Reflection
Imports System.Text
Imports Microsoft.Extensions.Logging
Imports ZoppaReports.Exceptions
Imports ZoppaReports.Tokens

''' <summary>環境変数クラス。</summary>
''' <typeparam name="T">対象クラス。</typeparam>
Public NotInheritable Class Environments

    ' パラメータ実体
    Private ReadOnly mTarget As Object

    ' プロパティディクショナリ
    Private ReadOnly mPropDic As New Dictionary(Of String, PropertyInfo)

    ' ローカル変数
    Private ReadOnly mVariants As New Dictionary(Of String, Object)

    ' リトライフラグ
    Private _retry As Boolean = False

    ' リトライ中ならば真
    Private Property Retrying As Boolean
        Get
            SyncLock Me
                Return Me._retry
            End SyncLock
        End Get
        Set(value As Boolean)
            SyncLock Me
                Me._retry = value
            End SyncLock
        End Set
    End Property

    ''' <summary>コンストラクタ。</summary>
    ''' <param name="type">パラメータの型。</param>
    ''' <param name="target">パラメータ。</param>
    ''' <summary>コンストラクタ。</summary>
    ''' <param name="target">パラメータ。</param>
    Public Sub New(target As Object)
        Me.mTarget = target
        If target IsNot Nothing Then
            Dim props = target.GetType().GetProperties()
            For Each prop In props
                Me.mPropDic.Add(prop.Name, prop)
            Next
        End If
        Me.Retrying = False
    End Sub

    ''' <summary>コピーコンストラクタ。</summary>
    ''' <param name="other">コピー元。</param>
    Private Sub New(other As Environments)
        Me.mTarget = other.mTarget
        Me.mPropDic = other.mPropDic
        Me.mVariants = New Dictionary(Of String, Object)(other.mVariants)
        Me.Retrying = other.Retrying
    End Sub

    ''' <summary>変数を消去します。</summary>
    Public Sub ClearVariant()
        Me.mVariants.Clear()
    End Sub

    ''' <summary>ローカル変数を追加する。</summary>
    ''' <param name="name">変数名。</param>
    ''' <param name="value">変数値。</param>
    Public Sub AddVariant(name As String, value As Object)
        If Me.mVariants.ContainsKey(name) Then
            Me.mVariants(name) = value
        Else
            Me.mVariants.Add(name, value)
        End If
    End Sub

    ''' <summary>指定した名称のプロパティから値を取得します。</summary>
    ''' <param name="name">プロパティ名。</param>
    ''' <returns>値。</returns>
    Public Function GetValue(name As String) As Object
        If Me.mVariants.ContainsKey(name) Then
            ' ローカル変数に存在しているため、それを返す
            Return Me.mVariants(name)

        ElseIf Not Me.mPropDic.ContainsKey(name) Then
            If Not Me.Retrying Then
                Try
                    Me.Retrying = True
                    Dim ans = ZoppaReports.ExecutesRefEnvironments(name, Me)
                    Return ans.Contents

                Catch ex As Exception
                    Logger.Value?.LogError("式を評価できません:{name}", name)
                Finally
                    Me.Retrying = False
                End Try
            End If
            Throw New ReportsAnalysisException($"指定したプロパティが定義されていません:{name}")
        Else
            ' 保持しているプロパティ参照から値を返す
            Return Me.mPropDic(name).GetValue(Me.mTarget)
        End If
    End Function

    ''' <summary>指定した名称のプロパティから値を取得します。</summary>
    ''' <param name="name">プロパティ名。</param>
    ''' <param name="index">インデックス。</param>
    ''' <returns>値。</returns>
    Public Function GetValue(name As String, index As Integer) As Object
        Dim v = Me.GetValue(name)
        If TypeOf v Is IEnumerable Then
            Dim i = 0
            For Each e In CType(v, IEnumerable)
                If i = index Then
                    Return e
                End If
                i += 1
            Next
            Throw New ReportsAnalysisException($"指定したインデックスが範囲外です:{index}")
        Else
            Throw New ReportsAnalysisException($"指定したプロパティは配列ではありません:{name}")
        End If
    End Function

    ''' <summary>指定した名称のメソッドを取得します。</summary>
    ''' <param name="name">メソッド名。</param>
    ''' <param name="value">値。</param>
    Public Function GetMethod(name As String) As Func(Of Object(), IToken)
        Select Case name.ToLower()
            Case "format"
                Return Function(args As Object())
                           If args.Length = 0 Then
                               Throw New ReportsAnalysisException("format関数に引数を設定してください。")
                           ElseIf args.Length = 1 Then
                               Return New StringToken(args(0).ToString(), """"c)
                           Else
                               Return New StringToken(String.Format(args(0).ToString(), args.Skip(1).ToArray()), """"c)
                           End If
                       End Function
        End Select
        Throw New ReportsAnalysisException("有効な関数を指定していません。")
    End Function

    ''' <summary>指定した名称のプロパティに値を設定します。</summary>
    ''' <param name="name">プロパティ名。</param>
    ''' <param name="value">値。</param>

    ''' <summary>環境値をコピーします。</summary>
    ''' <returns>コピーされた環境値。</returns>
    Public Function Clone() As Environments
        Return New Environments(Me)
    End Function

    ''' <summary>環境値をログ出力します。</summary>
    ''' <param name="logger">パラメータ。</param>
    Public Sub Logging(logger As ILogger)
        logger?.LogDebug("Parameters")
        For Each pair In Me.mPropDic
            Dim v = pair.Value.GetValue(Me.mTarget)
            logger?.LogDebug("・{pair.Key}={GetValueString(v)} ({pair.Value.PropertyType})", pair.Key, GetValueString(v), pair.Value.PropertyType)
        Next
    End Sub

    ''' <summary>オブジェクトの文字列表現を取得します。</summary>
    ''' <param name="value">オブジェクト。</param>
    ''' <returns>文字列表現。</returns>
    Private Shared Function GetValueString(value As Object) As String
        If TypeOf value Is IEnumerable Then
            Dim buf As New StringBuilder()
            For Each v In CType(value, IEnumerable)
                If buf.Length > 0 Then buf.Append(", ")
                buf.Append(v)
            Next
            Return buf.ToString()
        Else
            Return If(value?.ToString(), "[null]")
        End If
    End Function

End Class
