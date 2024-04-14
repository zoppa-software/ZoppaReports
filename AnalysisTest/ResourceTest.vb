Imports System.Drawing
Imports Xunit
Imports ZoppaReports
Imports ZoppaReports.Resources

Public Class ResourceTest

    <Fact>
    Public Sub BrushTest()
        Dim env = New With {.DefColor = Color.Red}

        Dim doc As New Xml.XmlDocument()
        doc.LoadXml("<brush name='red brush' color='#{DefColor}' />")
        Dim br = GetResourceFromXml(Of BrushResource)(doc.FirstChild, Nothing, env)
        Assert.Equal("red brush", br.Name)
        Assert.Equal(Color.Red, br.ForeColor)

        Dim doc2 As New Xml.XmlDocument()
        doc2.LoadXml("<brush name='hatch brush' color='Green' backcolor='White' hatchstyle='horizontal' />")
        Dim br2 = GetResourceFromXml(Of BrushResource)(doc2.FirstChild, Nothing)
        Assert.Equal("hatch brush", br2.Name)
        Assert.Equal(Color.Green, br2.ForeColor)
        Assert.Equal(Color.White, br2.BackColor)
        Assert.Equal(Drawing2D.HatchStyle.Horizontal, br2.HatchStyle)
    End Sub

    <Fact>
    Public Sub FontTest()
        Dim env = New With {.Fam = "Meiryo UI"}

        Dim doc As New Xml.XmlDocument()
        doc.LoadXml("<font name='def font' family='#{Fam}' emsize='9.0' style='italic' />")
        Dim fn = GetResourceFromXml(Of FontResource)(doc.FirstChild, Nothing, env)
        Assert.Equal("def font", fn.Name)
        Assert.Equal(9.0F, fn.EmSize)
        Assert.Equal(FontStyle.Italic, fn.Style)
    End Sub

    <Fact>
    Public Sub PenTest()
        Dim doc As New Xml.XmlDocument()
        doc.LoadXml("<pen name='def pen' color='255,0,0' width='1.5' />")
        Dim pn = GetResourceFromXml(Of PenResource)(doc.FirstChild, Nothing)
        Assert.Equal("def pen", pn.Name)
        Assert.Equal(Color.FromArgb(255, 0, 0), pn.Color)
        Assert.Equal(1.5F, pn.Width)
    End Sub

    <Fact>
    Public Sub ColorTest()
        Dim doc As New Xml.XmlDocument()
        doc.LoadXml("<color name='font color' color='Blue' />")
        Dim cr = GetResourceFromXml(Of ColorResource)(doc.FirstChild, Nothing)
        Assert.Equal("font color", cr.Name)
        Assert.Equal(Color.Blue, cr.Color)
    End Sub

    <Fact>
    Public Sub StyleTest()
        Dim doc As New Xml.XmlDocument()
        doc.LoadXml("<style name='style name'>
    <set name='Text' value='日本語' />
</style>")
        Dim st = GetResourceFromXml(Of StyleResource)(doc.FirstChild, Nothing)
        Assert.Equal("style name", st.Name)
    End Sub

End Class
