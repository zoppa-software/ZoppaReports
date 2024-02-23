Option Strict On
Option Explicit On

Imports System.Drawing
Imports System.Drawing.Printing
Imports System.Xml
Imports ZoppaReports.Exceptions
Imports ZoppaReports.Designer

''' <summary>レポート情報。</summary>
Public NotInheritable Class ReportsInformation

    ''' <summary>ページサイズを取得または設定します。</summary>
    Public Property Size As ReportsSize = ReportsSize.A4

    ''' <summary>設定文字列からレポート情報を読み込みます。</summary>
    ''' <param name="data">設定文字列。</param>
    ''' <param name="parameter">パラメータ。</param>
    ''' <returns>レポート情報。</returns>
    Friend Shared Function Load(data As String,
                                Optional parameter As Object = Nothing) As ReportsInformation
        ' XML情報から読み込む
        Dim doc = New Xml.XmlDocument()
        doc.LoadXml(data)

        ' XML情報を収集する
        Dim root = doc.SelectSingleNode("report")
        If root IsNot Nothing Then
            Dim result = New ReportsInformation()
            result.ReadNode(
                parameter,
                root,
                doc.SelectNodes("resources"),
                doc.SelectNodes("page")
            )
            Return result
        Else
            Throw New ReportsException("report要素が設定されていません")
        End If
    End Function

    ''' <summary>XMLノードから情報を読み込みます。</summary>
    ''' <param name="parameter">パラメータ。</param>
    ''' <param name="root">XMLノード。</param>
    ''' <param name="resources">リソースノード。</param>
    ''' <param name="pages">ページノード。</param>
    Private Sub ReadNode(parameter As Object,
                         root As XmlNode,
                         resources As XmlNodeList,
                         pages As XmlNodeList)
        Dim knd = root.Attributes("kind").GetValue(Of String)(parameter)
        Dim wid = root.Attributes("width").GetValue(Of Double)(parameter)
        Dim hig = root.Attributes("height").GetValue(Of Double)(parameter)
        Dim ort = root.Attributes("orientation").GetValue(Of PageOrientation)(parameter)

        ' サイズ情報を取得する
        Dim sz As ReportsSize = Nothing
        Dim ovl = If(ort.has, ort.value, PageOrientation.Landscape)

        If knd.has Then
            sz = ReportsSize.Search(knd.value)
            Me.Size = New ReportsSize(sz.Kind, sz.WidthInMM, sz.HeightInMM, ovl)
        ElseIf wid.has AndAlso hig.has Then
            Me.Size = New ReportsSize(PaperKind.Custom, wid.value, hig.value, ovl)
        Else
            Throw New ReportsException("有効なページ設定が行われていません")
        End If


    End Sub

    Public Sub Draw(g As Graphics)
        'Throw New NotImplementedException()
    End Sub



End Class
