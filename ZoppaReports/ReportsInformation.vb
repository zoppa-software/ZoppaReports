Option Strict On
Option Explicit On

Imports System.Drawing
Imports System.Drawing.Printing
Imports System.Xml
Imports ZoppaReports.Exceptions
Imports ZoppaReports.Designer
Imports ZoppaReports.Elements

''' <summary>レポート情報。</summary>
Public NotInheritable Class ReportsInformation
    Implements IReportsElements_old

#Region "properties"

    ''' <summary>ページサイズを取得または設定します。</summary>
    Public Property Size As ReportsSize = ReportsSize.A4

    ''' <summary>子要素リストを取得します。</summary>
    Public ReadOnly Property Elements As List(Of IReportsElements_old) = New List(Of IReportsElements_old)() Implements IReportsElements_old.Elements

    ''' <summary>パディングを取得または設定します。</summary>
    Public Property Padding As ReportsMargin = New ReportsMargin Implements IReportsElements_old.Padding

    ''' <summary>マージンを設定、取得します。</summary>
    ''' <returns>マージン情報。</returns>
    Public Property Margin As ReportsMargin Implements IReportsElements_old.Margin
        Get
            Return New ReportsMargin()
        End Get
        Set(value As ReportsMargin)
            Throw New ReportsException("マージンは変更できません")
        End Set
    End Property

    ''' <summary>左端位置を設定、取得します。</summary>
    ''' <returns>座標。</returns>
    Public Property LeftMM As Double Implements IReportsElements_old.LeftMM
        Get
            Return 0
        End Get
        Set(value As Double)
            Throw New ReportsException("左端は変更できません")
        End Set
    End Property

    ''' <summary>上端位置を設定、取得します。</summary>
    ''' <returns>座標。</returns>
    Public Property TopMM As Double Implements IReportsElements_old.TopMM
        Get
            Return 0
        End Get
        Set(value As Double)
            Throw New ReportsException("上端は変更できません")
        End Set
    End Property

    ''' <summary>幅を設定、取得します。</summary>
    ''' <returns>座標。</returns>
    Public Property WidthMM As Double Implements IReportsElements_old.WidthMM
        Get
            Return If(Me.Size.Orientation = ReportsOrientation_old.Portrait, Me.Size.WidthInMM, Me.Size.HeightInMM)
        End Get
        Set(value As Double)
            Throw New ReportsException("幅は変更できません")
        End Set
    End Property

    ''' <summary>高さを設定、取得します。</summary>
    ''' <returns>座標。</returns>
    Public Property HeightMM As Double Implements IReportsElements_old.HeightMM
        Get
            Return If(Me.Size.Orientation = ReportsOrientation_old.Portrait, Me.Size.HeightInMM, Me.Size.WidthInMM)
        End Get
        Set(value As Double)
            Throw New ReportsException("高さは変更できません")
        End Set
    End Property

#End Region

    ''' <summary>設定文字列からレポート情報を読み込みます。</summary>
    ''' <param name="data">設定文字列。</param>
    ''' <returns>レポート情報。</returns>
    Friend Shared Function Load(data As String) As ReportsInformation
        ' XML情報から読み込む
        Dim doc = New Xml.XmlDocument()
        doc.LoadXml(data)

        ' XML情報を収集する
        Dim root = doc.SelectSingleNode("report")
        If root IsNot Nothing Then
            Dim result = New ReportsInformation()
            result.ReadNode(
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
    ''' <param name="root">XMLノード。</param>
    ''' <param name="resources">リソースノード。</param>
    ''' <param name="pages">ページノード。</param>
    Private Sub ReadNode(root As XmlNode,
                         resources As XmlNodeList,
                         pages As XmlNodeList)
        Dim knd = root.Attributes("kind").GetValue(Of ReportsSize)()
        Dim wid = root.Attributes("width").GetValueDouble()
        Dim hig = root.Attributes("height").GetValueDouble()
        Dim ort = root.Attributes("orientation").GetValue(Of ReportsOrientation_old)()

        ' サイズ情報を取得する
        Dim ovl = If(ort.has, ort.value, knd.value.Orientation)
        If knd.has Then
            Me.Size = New ReportsSize(knd.value.Kind, knd.value.WidthInMM, knd.value.HeightInMM, ovl)
        ElseIf wid.has AndAlso hig.has Then
            Me.Size = New ReportsSize(PaperKind.Custom, wid.value, hig.value, ovl)
        Else
            Throw New ReportsException("有効なページ設定が行われていません")
        End If

        ' パディング情報を取得する
        Dim pad = root.Attributes("padding").GetValue(Of ReportsMargin)()
        If pad.has Then
            Me.Padding = pad.value
        End If
    End Sub



    Public Sub CalcLocation(Optional parameter As Object = Nothing)
        Throw New NotImplementedException()
    End Sub

    Public Sub Draw(g As Graphics)
        g.PageUnit = GraphicsUnit.Point
        ''Throw New NotImplementedException()

        g.DrawLine(Pens.Black, 0, 14.173, 100, 14.173)
        g.DrawLine(Pens.Black, 0, 28.346, 100, 28.346)

        g.DrawString("あいう", New Font("MS Gothic", 12), Brushes.Black, 0, 0)

        g.Flush()
    End Sub

End Class
