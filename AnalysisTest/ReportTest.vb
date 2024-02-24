Imports System.Drawing
Imports System.Drawing.Printing
Imports System.Printing
Imports Xunit
Imports ZoppaReports
Imports ZoppaReports.Designer
Imports ZoppaReports.Exceptions

Public Class ReportTest

    <Fact>
    Public Sub PageTest()
        Dim pageXml =
"<report kind=""A4"">
    <resources>
        <font name=""MS Gothic"" size=""12"" />
    </resources>
    <page>
    </page>
</report>
"
        Dim a1 = ReadReportsInformation(pageXml)
        Assert.Equal(PaperKind.A4, a1.Size.Kind)

        Dim pageXml2 =
"<report kind=""#{kind}"">
    <page>
    </page>
</report>
"
        Dim a2 = ReadReportsInformation(pageXml2, New With {.kind = ReportsSize.B4})
        Assert.Equal(PaperKind.B4, a2.Size.Kind)

        Dim pageXml3 =
"<report width=""100"" height=""200"" orientation=""portrait"">
    <page>
    </page>
</report>
"
        Dim a3 = ReadReportsInformation(pageXml3)
        Assert.Equal(PaperKind.Custom, a3.Size.Kind)
        Assert.Equal(100, a3.Size.WidthInMM)
        Assert.Equal(200, a3.Size.HeightInMM)
        Assert.Equal(ReportsOrientation.Portrait, a3.Size.Orientation)

        Dim pageXml4 =
"<report kind=""custom,140,200,portrait"">
    <page>
    </page>
</report>
"
        Dim a4 = ReadReportsInformation(pageXml4)
        Assert.Equal(PaperKind.Custom, a4.Size.Kind)
        Assert.Equal(140, a4.Size.WidthInMM)
        Assert.Equal(200, a4.Size.HeightInMM)
        Assert.Equal(ReportsOrientation.Portrait, a4.Size.Orientation)

        Dim pageXml5 =
"<report width=""#{wid + 10}"" height=""#{hid - 10}"" orientation=""#{ori}"">
    <page>
    </page>
</report>
"
        Dim a5 = ReadReportsInformation(pageXml5, New With {.wid = 90, .hid = 110, .ori = ReportsOrientation.Landscape})
        Assert.Equal(PaperKind.Custom, a5.Size.Kind)
        Assert.Equal(100, a5.Size.WidthInMM)
        Assert.Equal(100, a5.Size.HeightInMM)
        Assert.Equal(ReportsOrientation.Landscape, a5.Size.Orientation)
    End Sub

    <Fact>
    Public Sub PageErrorTest()
        Dim pageXml =
"<reporterr kind=""A4"">
    <resources>
        <font name=""MS Gothic"" size=""12"" />
    </resources>
    <page>
    </page>
</reporterr>
"
        ' report要素未定義
        Assert.Throws(Of ReportsException)(
            Sub()
                ReadReportsInformation(pageXml)
            End Sub
        )
    End Sub

    <Fact>
    Public Sub MarginTest()
        Dim pageXml =
"<report kind=""A4"" padding=""3"">
    <page>
    </page>
</report>
"
        Dim a1 = ReadReportsInformation(pageXml)
        Assert.Equal(New ReportsMargin With {.Left = 3, .Top = 3, .Right = 3, .Bottom = 3}, a1.Padding)

        Dim pageXml1 =
"<report kind=""A4"" padding=""3, 5"">
    <page>
    </page>
</report>
"
        Dim a2 = ReadReportsInformation(pageXml1)
        Assert.Equal(New ReportsMargin With {.Left = 3, .Top = 5, .Right = 3, .Bottom = 5}, a2.Padding)

        Dim pageXml2 =
"<report kind=""A4"" padding=""1, 2, 3, 4"">
    <page>
    </page>
</report>
"
        Dim a3 = ReadReportsInformation(pageXml2)
        Assert.Equal(New ReportsMargin With {.Left = 1, .Top = 2, .Right = 3, .Bottom = 4}, a3.Padding)
    End Sub

End Class
