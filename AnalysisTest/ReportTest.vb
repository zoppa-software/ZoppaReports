Imports System.Drawing
Imports System.Drawing.Printing
Imports Xunit
Imports ZoppaReports
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
        Dim info = ReadReportsInformation(pageXml)
        Assert.Equal(PaperKind.A4, info.Size.Kind)
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
