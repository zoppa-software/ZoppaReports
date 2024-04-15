Imports System.Drawing.Printing
Imports Microsoft.Extensions.Logging
Imports Xunit
Imports ZoppaReports
Imports ZoppaReports.Smalls

Public Class ReportElementTest

    <Fact>
    Public Sub ReportdTest()
        Using logFacyory = LoggerFactory.Create(
                Sub(config)
                    config.SetMinimumLevel(LogLevel.Trace)
                    config.AddConsole()
                End Sub)
            SetZoppaReportsLogFactory(logFacyory)

            Dim env = New With {.PageSize = ReportsSize.A3}

            Dim pageXml1 =
"<report size='#{PageSize}'>
    <resources>
        <font name='lb font' family='Meiryo UI' emsize='9.0' style='Italic' />
        <style base='label'>
            <set name='font' value='@{lb font}' />
        </style>
    </resources>
    <label text='123' />
</report>"
            Dim ans1 = CType(CreateReportsStructFromXml(pageXml1, env).info, ReportsElement)
            Assert.Equal(ReportsSize.A3, ans1.PaperSize)
            Assert.Equal(ReportsOrientation.Portrait, ans1.PaperOrientation)
        End Using
    End Sub

End Class
