Option Strict On
Option Explicit On

Imports System.Drawing
Imports System.Runtime.CompilerServices
Imports System.Text.RegularExpressions
Imports ZoppaReports.Resources
Imports ZoppaReports.Smalls

''' <summary>リーダー補助モジュール。</summary>
Public Module ReaderHelper

    ' 色パターンの正規表現
    Private ReadOnly mColorPtnRegex As New Regex("^[0-9]{1,3}(,\s*[0-9]{1,3}){2,3}$", RegexOptions.Compiled)

    ' デフォルトフォント
    Private ReadOnly mDefaultFont As New Font("Meiryo", 12)

    ''' <summary>デフォルトフォントを取得します。</summary>
    Public ReadOnly Property DefaultFont As Font
        Get
            Return mDefaultFont
        End Get
    End Property

    ''' <summary>ノードを取得します。</summary>
    ''' <param name="nodeList">ノードリスト。</param>
    ''' <param name="nodeName">ノード名。</param>
    ''' <returns>ノード。</returns>
    <Extension()>
    Public Function GetNodeByName(nodeList As Xml.XmlNodeList, nodeName As String) As Xml.XmlNode
        For Each nd As Xml.XmlNode In nodeList
            If nd.Name.ToLower() = nodeName.ToLower() Then
                Return nd
            End If
        Next
        Return Nothing
    End Function

    ''' <summary>属性を取得します。</summary>
    ''' <param name="attrList">属性リスト。</param>
    ''' <param name="attrName">属性名。</param>
    ''' <returns>属性。</returns>
    <Extension()>
    Public Function GetAttributeByName(attrList As Xml.XmlAttributeCollection, attrName As String) As Xml.XmlAttribute
        For Each nd As Xml.XmlAttribute In attrList
            If nd.Name.ToLower() = attrName.ToLower() Then
                Return nd
            End If
        Next
        Return Nothing
    End Function

    ''' <summary>値を置き換えます。</summary>
    ''' <param name="value">置き換え対象の値。</param>
    ''' <param name="env">バインディングデータ。</param>
    ''' <param name="element">要素。</param>
    ''' <returns>置き換え後の値。</returns>
    Friend Function Replace(value As String, env As Environments, element As IReportsElement) As Object
        Dim ret As Object = value
        If value.StartsWith("#{") AndAlso value.EndsWith("}") Then
            ' 置き換え式なら置き換え
            Dim expr = value.Substring(2, value.Length - 3)
            ret = ExecutesRefEnvironments(expr, env).Contents

        ElseIf value.StartsWith("@{") AndAlso value.EndsWith("}") AndAlso element IsNot Nothing Then
            ' リソース参照ならリソースを取得
            Dim expr = value.Substring(2, value.Length - 3).Trim()
            ret = element.GetResource(Of IReportsResources)(expr).Contents(Of Object)()
        End If
        Return ret
    End Function

    ''' <summary>色を取得します。</summary>
    ''' <param name="value">XMLテンプレートから取得した値。</param>
    ''' <returns>色。</returns>
    Friend Function ConvertColor(value As Object) As Color
        If value.GetType() Is GetType(Color) Then
            ' 色を返す
            Return CType(CObj(value), Color)

        ElseIf mColorPtnRegex.IsMatch(value.ToString()) Then
            Try
                Dim arr = value.ToString().Split(","c)
                If arr.Length = 3 Then
                    Return Color.FromArgb(CInt(arr.GetValue(0)), CInt(arr.GetValue(1)), CInt(arr.GetValue(2)))
                ElseIf arr.Length = 4 Then
                    Return Color.FromArgb(CInt(arr.GetValue(0)), CInt(arr.GetValue(1)), CInt(arr.GetValue(2)), CInt(arr.GetValue(3)))
                End If
            Catch ex As Exception
                ' 例外は無視
            End Try
        Else
            ' 色名から色を取得
            Dim kColor As KnownColor
            If [Enum].TryParse(Of KnownColor)(value.ToString(), True, kColor) Then
                Return Color.FromKnownColor(kColor)
            End If
        End If
        Return Color.Transparent
    End Function

    ''' <summary>フォントを取得します。</summary>
    ''' <param name="value">XMLテンプレートから取得した値。</param>
    ''' <returns>フォント。</returns>
    Friend Function ConvertFont(value As Object) As Font
        If value.GetType() Is GetType(Font) Then
            Return CType(CObj(value), Font)
        Else
            Return DefaultFont
        End If
    End Function

    ''' <summary>ブラシを取得します。</summary>
    ''' <param name="value">XMLテンプレートから取得した値。</param>
    ''' <returns>ブラシ。</returns>
    Friend Function ConvertBrush(value As Object) As Brush
        If value.GetType() Is GetType(Brush) Then
            ' ブラシを返す
            Return CType(CObj(value), Brush)

        Else
            ' ブラシ名からブラシを取得
            For Each bprop In GetType(Brushes).GetProperties()
                If bprop.Name.ToLower() = value.ToString().ToLower() Then
                    Return CType(bprop.GetValue(Nothing), Brush)
                End If
            Next
        End If
        Return Brushes.Black
    End Function

    Friend Function ConvertScalar(value As Object) As ReportsScalar
        If value.GetType() Is GetType(ReportsScalar) Then
            Return CType(CObj(value), ReportsScalar)
        Else
            Return CType(value.ToString(), ReportsScalar)
        End If
    End Function

End Module
