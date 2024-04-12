Imports System.Drawing
Imports Xunit
Imports ZoppaReports.Designer
Imports ZoppaReports.Resources

Public Class ResourceTest

    <Fact>
    Public Sub BrushTest()
        Dim env = New With {.DefColor = Color.Red}

        Dim doc As New Xml.XmlDocument()
        doc.LoadXml("<brush name='red brush' color='#{DefColor}' />")
        Dim br = GetResourceFromXml(Of BrushResource)(doc.FirstChild, env)
        Assert.Equal("red brush", br.Name)
        Assert.Equal(Color.Red, br.ForeColor)

        Dim doc2 As New Xml.XmlDocument()
        doc2.LoadXml("<brush name='hatch brush' color='Green' backcolor='White' hatchstyle='Horizontal' />")
        Dim br2 = GetResourceFromXml(Of BrushResource)(doc2.FirstChild)
        Assert.Equal("hatch brush", br2.Name)
        Assert.Equal(Color.Green, br2.ForeColor)
        Assert.Equal(Color.White, br2.BackColor)
        Assert.Equal(Drawing2D.HatchStyle.Horizontal, br2.HatchStyle)
    End Sub

    <Fact>
    Public Sub FontTest()
        Dim env = New With {.Fam = "Meiryo UI"}

        Dim doc As New Xml.XmlDocument()
        doc.LoadXml("<font name='def font' family='#{Fam}' emsize='9.0' style='Italic' />")
        Dim fn = GetResourceFromXml(Of FontResource)(doc.FirstChild, env)
        Assert.Equal("def font", fn.Name)
        Assert.Equal(9.0F, fn.EmSize)
        Assert.Equal(FontStyle.Italic, fn.Style)
    End Sub

    <Fact>
    Public Sub PenTest()
        Dim doc As New Xml.XmlDocument()
        doc.LoadXml("<pen name='def pen' color='#{[255,0,0]}' width='1.5' />")
        Dim pn = GetResourceFromXml(Of PenResource)(doc.FirstChild)
        Assert.Equal("def pen", pn.Name)
        Assert.Equal(Color.FromArgb(255, 0, 0), pn.Color)
        Assert.Equal(1.5F, pn.Width)
    End Sub

    <Fact>
    Public Sub ColorTest()
        Dim doc As New Xml.XmlDocument()
        doc.LoadXml("<color name='font color' color='Blue' />")
        Dim cr = GetResourceFromXml(Of ColorResource)(doc.FirstChild)
        Assert.Equal("font color", cr.Name)
        Assert.Equal(Color.Blue, cr.Color)
    End Sub


    '<Fact>
    'Public Sub PageOrientationTest()
    '    Dim ai = CType(ReportsOrientation.Landscape, String)
    '    Assert.Equal("landscape", ai)

    '    Dim ai2 = CType(ReportsOrientation.Portrait, String)
    '    Assert.Equal("portrait", ai2)

    '    Dim ai3 = CType("landscape", ReportsOrientation)
    '    Assert.Equal(ReportsOrientation.Landscape, ai3)

    '    Dim ai4 = CType("portrait", ReportsOrientation)
    '    Assert.Equal(ReportsOrientation.Portrait, ai4)

    '    Dim a1 = (ReportsOrientation.Landscape = ReportsOrientation.Landscape)
    '    Assert.True(a1)

    '    Dim a11 = (ReportsOrientation.Landscape = Nothing)
    '    Assert.False(a11)

    '    Dim a12 = (Nothing = ReportsOrientation.Landscape)
    '    Assert.False(a12)

    '    Dim a2 = (ReportsOrientation.Landscape <> ReportsOrientation.Portrait)
    '    Assert.True(a2)

    '    Dim a21 = ReportsOrientation.Landscape.Equals(ReportsOrientation.Landscape)
    '    Assert.True(a21)

    '    Dim hc = ReportsOrientation.Landscape.GetHashCode()

    '    Assert.Throws(Of InvalidCastException)(Function() CType("xxx", ReportsOrientation))
    'End Sub

    '<Fact>
    'Public Sub ReportsSizeTest()

    'End Sub

    '<Fact>
    'Public Sub Ooo()
    '    Dim o As New FontResource("aaa")
    '    o.SetProperty(NameOf(o.FamilyName), "f_name")
    '    o.SetProperty(NameOf(o.EmSize), "9.0")
    '    o.SetProperty(NameOf(o.Style), "Bold")

    '    Dim o2 As New PenResource("bbb")
    '    o2.SetProperty(NameOf(o2.Color), Color.Red)
    '    o2.SetProperty(NameOf(o2.Width), "2.0")
    'End Sub

End Class
