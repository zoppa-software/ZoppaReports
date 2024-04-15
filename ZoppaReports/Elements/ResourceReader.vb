Option Strict On
Option Explicit On

Imports System.Drawing
Imports System.Windows
Imports ZoppaReports.Resources

''' <summary>リソースリーダー。</summary>
Public Module ResourceReader

    ''' <summary>XMLノードからリソースを取得します。</summary>
    ''' <typeparam name="T">リソース型。</typeparam>
    ''' <param name="node">XMLノード。</param>
    ''' <param name="element">要素。</param>
    ''' <param name="bindingData">バインディングデータ。</param>
    ''' <returns>リソース。</returns>
    Public Function GetResourceFromXml(Of T As IReportsResources)(node As Xml.XmlNode,
                                                                  element As IReportsElement,
                                                                  Optional bindingData As Object = Nothing) As T
        Return CType(LocalResourceFromXml(node, element, New Environments(bindingData)), T)
    End Function

    ''' <summary>XMLノードからリソースを取得します。</summary>
    ''' <param name="node">XMLノード。</param>
    ''' <param name="element">要素。</param>
    ''' <param name="env">バインディングデータ。</param>
    ''' <returns>リソース。</returns>
    Public Function GetResourceFromXml(node As Xml.XmlNode,
                                       element As IReportsElement,
                                       env As Environments) As IReportsResources
        Return LocalResourceFromXml(node, element, env)
    End Function

    ''' <summary>XMLノードからリソースを取得します。</summary>
    ''' <param name="node">XMLノード。</param>
    ''' <param name="element">要素。</param>
    ''' <param name="env">バインディングデータ。</param>
    ''' <returns>リソース。</returns>
    Private Function LocalResourceFromXml(node As Xml.XmlNode,
                                          element As IReportsElement,
                                          env As Environments) As IReportsResources
        Dim nm = node.Attributes.GetAttributeByName("name")

        Select Case node.Name.ToLower()
            Case "brush"
                Dim bres As New BrushResource(nm.Value)
                SetResProperties(bres, node, element, env)
                Return bres

            Case "font"
                Dim fres As New FontResource(nm.Value)
                SetResProperties(fres, node, element, env)
                Return fres

            Case "pen"
                Dim pres As New PenResource(nm.Value)
                SetResProperties(pres, node, element, env)
                Return pres

            Case "color"
                Dim cres As New ColorResource(nm.Value)
                SetResProperties(cres, node, element, env)
                Return cres

            Case "style"
                Dim sres As New StyleResource(nm?.Value)
                SetResProperties(sres, node, element, env)
                For Each nd As Xml.XmlNode In node.ChildNodes
                    sres.AddSetter(ConvertSetter(nd, env, element))
                Next
                Return sres

            Case Else
                Return Nothing
        End Select
    End Function

    ''' <summary>リソースのプロパティを設定します。</summary>
    ''' <param name="ansres">リソース。</param>
    ''' <param name="node">XMLノード。</param>
    ''' <param name="element">要素。</param>
    ''' <param name="env">バインディングデータ。</param>
    Private Sub SetResProperties(ansres As IReportsResources, node As Xml.XmlNode, element As IReportsElement, env As Environments)
        For Each attr As Xml.XmlAttribute In node.Attributes
            ansres.SetProperty(attr.Name, Replace(attr.Value, env, element))
        Next
    End Sub

    ''' <summary>セッターを変換します。</summary>
    ''' <param name="nd">XMLノード。</param>
    ''' <param name="env">バインディングデータ。</param>
    ''' <param name="element">要素。</param>
    ''' <returns>セッター。</returns>
    Friend Function ConvertSetter(nd As Xml.XmlNode,
                                  env As Environments,
                                  element As IReportsElement) As (key As String, val As Object)
        If nd.Name.ToLower() = "set" Then
            ' 名前
            Dim nm = nd.Attributes.GetAttributeByName("name")

            ' 値
            Dim vl = nd.Attributes.GetAttributeByName("value")

            If nm IsNot Nothing AndAlso vl IsNot Nothing Then
                Return (nm.Value, Replace(vl.Value, env, element))
            End If
        End If
        Return ("", Nothing)
    End Function

End Module


