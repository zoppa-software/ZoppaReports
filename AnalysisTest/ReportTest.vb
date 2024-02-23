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
"<report width=""#{wid + 10}"" height=""#{hid - 10}"" orientation=""#{ori}"">
    <page>
    </page>
</report>
"
        Dim a4 = ReadReportsInformation(pageXml4, New With {.wid = 90, .hid = 110, .ori = ReportsOrientation.Landscape})
        Assert.Equal(PaperKind.Custom, a4.Size.Kind)
        Assert.Equal(100, a4.Size.WidthInMM)
        Assert.Equal(100, a4.Size.HeightInMM)
        Assert.Equal(ReportsOrientation.Landscape, a4.Size.Orientation)
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

End Class
