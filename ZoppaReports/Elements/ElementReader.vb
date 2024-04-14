Option Strict On
Option Explicit On

Imports ZoppaReports.Draw
Imports ZoppaReports.Resources

''' <summary>要素リーダー。</summary>
Public Module ElementReader

    Public Const RESOURCES_TAG As String = "resources"

    Public Const LABEL_TAG As String = "label"

    ''' <summary>XMLノードから要素を取得します。</summary>
    ''' <typeparam name="T">要素型。</typeparam>
    ''' <param name="node">XMLノード。</param>
    ''' <param name="element">要素。</param>
    ''' <param name="bindingData">バインディングデータ。</param>
    ''' <returns>リソース。</returns>
    Public Function GetElementFromXml(Of T As IReportsElement)(node As Xml.XmlNode,
                                                               element As IReportsElement,
                                                               Optional bindingData As Object = Nothing) As T
        Return CType(LocalElementFromXml(node, element, New Environments(bindingData)), T)
    End Function

    ''' <summary>XMLノードから要素を取得します。</summary>
    ''' <param name="node">XMLノード。</param>
    ''' <param name="element">要素。</param>
    ''' <param name="env">バインディングデータ。</param>
    ''' <returns>要素。</returns>
    Public Function GetElementFromXml(node As Xml.XmlNode,
                                      element As IReportsElement,
                                      env As Environments) As IReportsElement
        Return LocalElementFromXml(node, element, env)
    End Function

    ''' <summary>XMLノードから要素を取得します。</summary>
    ''' <param name="node">XMLノード。</param>
    ''' <param name="element">要素。</param>
    ''' <param name="env">バインディングデータ。</param>
    ''' <returns>要素。</returns>
    Private Function LocalElementFromXml(node As Xml.XmlNode,
                                         element As IReportsElement,
                                         env As Environments) As IReportsElement
        Dim res As IReportsElement

        Select Case node.Name.ToLower()
            Case "report"
                Dim rep As New ReportsElement()
                CollectResources(node, rep, env)
                rep.GetBaseStyle(GetType(ReportsElement))?.ExpandProperties(rep)
                For Each attr As Xml.XmlAttribute In node.Attributes
                    rep.SetProperty(attr.Name, Replace(attr.Value, env, element))
                Next
                res = rep

            Case LABEL_TAG
                Dim lele As New LabelElement(element)
                CollectResources(node, lele, env)
                lele.GetBaseStyle(GetType(LabelElement))?.ExpandProperties(lele)
                For Each attr As Xml.XmlAttribute In node.Attributes
                    lele.SetProperty(attr.Name, Replace(attr.Value, env, element))
                Next
                res = lele

            Case Else
                res = Nothing
        End Select

        '
        For Each nd As Xml.XmlNode In node.ChildNodes
            If nd.Name.ToLower() <> RESOURCES_TAG Then
                Dim ch = LocalElementFromXml(nd, res, env)
                If ch IsNot Nothing Then
                    res.AddChildElement(ch)
                End If
            End If
        Next

        Return res
    End Function

    Public Function GetElementType(tagName As String) As Type
        Select Case tagName.ToLower()
            Case LABEL_TAG
                Return GetType(LabelElement)

            Case Else
                Return Nothing
        End Select
    End Function

    ''' <summary>XMLノードからリソースを収集します。</summary>
    ''' <param name="node">XMLノード。</param>
    ''' <param name="element">要素。</param>
    ''' <param name="bindingData">バインディングデータ。</param>
    ''' <returns>リソースリスト。</returns>
    Private Sub CollectResources(node As Xml.XmlNode, element As IReportsElement, env As Environments)
        ' リソースノードを取得
        Dim resNode = node.ChildNodes.GetNodeByName(RESOURCES_TAG)

        ' リソースを追加
        If resNode IsNot Nothing Then
            For Each nd As Xml.XmlNode In resNode.ChildNodes
                Dim r = GetResourceFromXml(nd, element, env)
                element.AddResources(r)
            Next
        End If
    End Sub

End Module


