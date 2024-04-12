Option Strict On
Option Explicit On

Imports System.Drawing

Namespace Resources

    ''' <summary>リソースリーダー。</summary>
    Public Module ResourceReader

        ' 生成したリソースリスト
        Private mAllResources As New SortedList(Of String, IResources)

        Public Sub ClearResources()
            mAllResources.Clear()
        End Sub

        Public Function GetResourceFromXml(Of T As IResources)(node As Xml.XmlNode, Optional bindingData As Object = Nothing) As T
            Dim res = LocalResourceFromXml(node, bindingData)
            Return CType(res, T)
        End Function

        Private Function LocalResourceFromXml(node As Xml.XmlNode, bindingData As Object) As IResources
            Dim nm = node.Attributes("name")
            If nm Is Nothing Then
                nm = node.Attributes("Name")
            End If

            Select Case node.Name.ToLower()
                Case "brush"
                    Dim bres As New BrushResource(nm.Value)
                    'baseStyle?.ExpandProperties(bres)
                    'bres.Style?.ExpandProperties(bres)

                    For Each attr As Xml.XmlAttribute In node.Attributes
                        bres.SetProperty(attr.Name, Replace(attr.Value, bindingData))
                    Next
                    Return bres

                Case "font"
                    Dim fres As New FontResource(nm.Value)
                    'baseStyle?.ExpandProperties(bres)
                    'bres.Style?.ExpandProperties(bres)

                    For Each attr As Xml.XmlAttribute In node.Attributes
                        fres.SetProperty(attr.Name, Replace(attr.Value, bindingData))
                    Next
                    Return fres

                Case "pen"
                    Dim pres As New PenResource(nm.Value)
                    'baseStyle?.ExpandProperties(bres)
                    'bres.Style?.ExpandProperties(bres)

                    For Each attr As Xml.XmlAttribute In node.Attributes
                        pres.SetProperty(attr.Name, Replace(attr.Value, bindingData))
                    Next
                    Return pres

                Case "color"
                    Dim cres As New ColorResource(nm.Value)
                    'baseStyle?.ExpandProperties(bres)
                    'bres.Style?.ExpandProperties(bres)

                    For Each attr As Xml.XmlAttribute In node.Attributes
                        cres.SetProperty(attr.Name, Replace(attr.Value, bindingData))
                    Next
                    Return cres

                Case Else
                    Return Nothing
            End Select
        End Function

        Private Function Replace(value As String, bindingData As Object) As Object
            Dim ret As Object = value
            If value.StartsWith("#{") AndAlso value.EndsWith("}") Then
                Dim expr = value.Substring(2, value.Length - 3)
                ret = expr.Executes(bindingData).Contents

            ElseIf value.StartsWith("@{") AndAlso value.EndsWith("}") Then
                Dim expr = value.Substring(2, value.Length - 3).Trim()
                If mAllResources.ContainsKey(expr) Then
                    ret = mAllResources(expr).Contents(Of Object)()
                End If
            End If
            Return ret
        End Function

        Friend Function ConvertColor(value As Object) As Color
            If value.GetType() Is GetType(Color) Then
                Return CType(CObj(value), Color)

            ElseIf value.GetType().IsArray Then
                Try
                    Dim arr = CType(value, Object())
                    If arr.Length = 3 Then
                        Return Color.FromArgb(CInt(arr.GetValue(0)), CInt(arr.GetValue(1)), CInt(arr.GetValue(2)))
                    ElseIf arr.Length = 4 Then
                        Return Color.FromArgb(CInt(arr.GetValue(0)), CInt(arr.GetValue(1)), CInt(arr.GetValue(2)), CInt(arr.GetValue(3)))
                    Else
                        Return Color.Transparent
                    End If
                Catch ex As Exception
                    Return Color.Transparent
                End Try
            Else
                Return Color.FromName(value.ToString())
            End If
        End Function

    End Module

End Namespace

