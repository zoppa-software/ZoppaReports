Imports System.Drawing
Imports System.Drawing.Printing
Imports System.Printing
Imports Microsoft.Extensions.Logging
Imports Xunit
Imports ZoppaReports
Imports ZoppaReports.Designer
Imports ZoppaReports.Exceptions

Public Class ReportTest

    <Fact>
    Public Sub PageTest()
        Using logFacyory = LoggerFactory.Create(
                Sub(config)
                    config.SetMinimumLevel(LogLevel.Trace)
                    config.AddConsole()
                End Sub)
            SetZoppaReportsLogFactory(logFacyory)
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
            Assert.Equal(PaperKind.A4, a1.info.Size.Kind)

            Dim pageXml2 =
"<report kind=""#{kind}"">
    <page>
    </page>
</report>
"
            Dim a2 = ReadReportsInformation(pageXml2, New With {.kind = ReportsSize.B4})
            Assert.Equal(PaperKind.B4, a2.info.Size.Kind)

            Dim pageXml3 =
"<report width=""100"" height=""200"" orientation=""portrait"">
    <page>
    </page>
</report>
"
            Dim a3 = ReadReportsInformation(pageXml3)
            Assert.Equal(PaperKind.Custom, a3.info.Size.Kind)
            Assert.Equal(100, a3.info.Size.WidthInMM)
            Assert.Equal(200, a3.info.Size.HeightInMM)
            Assert.Equal(ReportsOrientation.Portrait, a3.info.Size.Orientation)

            Dim pageXml4 =
"<report kind=""custom,140,200,portrait"">
    <page>
    </page>
</report>
"
            Dim a4 = ReadReportsInformation(pageXml4)
            Assert.Equal(PaperKind.Custom, a4.info.Size.Kind)
            Assert.Equal(140, a4.info.Size.WidthInMM)
            Assert.Equal(200, a4.info.Size.HeightInMM)
            Assert.Equal(ReportsOrientation.Portrait, a4.info.Size.Orientation)

            Dim pageXml5 =
"<report width=""#{wid + 10}"" height=""#{hid - 10}"" orientation=""#{ori}"">
    <page>
    </page>
</report>
"
            Dim a5 = ReadReportsInformation(pageXml5, New With {.wid = 90, .hid = 110, .ori = ReportsOrientation.Landscape})
            Assert.Equal(PaperKind.Custom, a5.info.Size.Kind)
            Assert.Equal(100, a5.info.Size.WidthInMM)
            Assert.Equal(100, a5.info.Size.HeightInMM)
            Assert.Equal(ReportsOrientation.Landscape, a5.info.Size.Orientation)
        End Using
    End Sub

    <Fact>
    Public Sub PageErrorTest()
        Using logFacyory = LoggerFactory.Create(
                Sub(config)
                    config.SetMinimumLevel(LogLevel.Trace)
                    config.AddConsole()
                End Sub)
            SetZoppaReportsLogFactory(logFacyory)
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
        End Using
    End Sub

    <Fact>
    Public Sub MarginTest()
        Using logFacyory = LoggerFactory.Create(
                Sub(config)
                    config.SetMinimumLevel(LogLevel.Trace)
                    config.AddConsole()
                End Sub)
            SetZoppaReportsLogFactory(logFacyory)
            Dim pageXml =
"<report kind=""A4"" padding=""3"">
    <page>
    </page>
</report>
"
            Dim a1 = ReadReportsInformation(pageXml)
            Assert.Equal(New ReportsMargin With {.Left = 3, .Top = 3, .Right = 3, .Bottom = 3}, a1.info.Padding)

            Dim pageXml1 =
    "<report kind=""A4"" padding=""3, 5"">
    <page>
    </page>
</report>
"
            Dim a2 = ReadReportsInformation(pageXml1)
            Assert.Equal(New ReportsMargin With {.Left = 3, .Top = 5, .Right = 3, .Bottom = 5}, a2.info.Padding)

            Dim pageXml2 =
    "<report kind=""A4"" padding=""1, 2, 3, 4"">
    <page>
    </page>
</report>
"
            Dim a3 = ReadReportsInformation(pageXml2)
            Assert.Equal(New ReportsMargin With {.Left = 1, .Top = 2, .Right = 3, .Bottom = 4}, a3.info.Padding)
        End Using
    End Sub

    <Fact>
    Public Sub IfTest()
        Dim pageXml1 =
"<report kind=""A4"" padding=""5"">
    <page>
        {if no < 10}
        <label x=""10"" y=""20"" width=""30"" height=""6"" brush=""black""/>
        {elseif no < 20}
        <label x=""40"" y=""50"" width=""60"" height=""12"" brush=""black""/>
        {else}
        <label x=""70"" y=""80"" width=""90"" height=""18"" brush=""black""/>
        {end if}
    </page>
</report>"
        Dim a1 = ReadReportsInformation(pageXml1, New With {.no = 5})
        Assert.Equal("<report kind=""A4"" padding=""5"">
    <page>
        
        <label x=""10"" y=""20"" width=""30"" height=""6"" brush=""black""/>
        
    </page>
</report>", a1.repdata)

        Dim a2 = ReadReportsInformation(pageXml1, New With {.no = 19})
        Assert.Equal("<report kind=""A4"" padding=""5"">
    <page>
        
        <label x=""40"" y=""50"" width=""60"" height=""12"" brush=""black""/>
        
    </page>
</report>", a2.repdata)

        Dim a3 = ReadReportsInformation(pageXml1, New With {.no = 30})
        Assert.Equal("<report kind=""A4"" padding=""5"">
    <page>
        
        <label x=""70"" y=""80"" width=""90"" height=""18"" brush=""black""/>
        
    </page>
</report>", a3.repdata)
    End Sub

    <Fact>
    Public Sub ForTest()
        Using logFacyory = LoggerFactory.Create(
            Sub(config)
                config.SetMinimumLevel(LogLevel.Trace)
                config.AddConsole()
            End Sub)
            SetZoppaReportsLogFactory(logFacyory)

            Dim pageXml1 =
"<report kind=""A4"" padding=""5"">
    <page>
        {for i = base + 0 to base + 9}
        <label x=""0"" y=""#{i * 5}"" width=""100"" height=""5"" brush=""black""/>{/for}
    </page>
</report>"
            Dim a1 = ReadReportsInformation(pageXml1, New With {.base = 2})
            Assert.Equal("<report kind=""A4"" padding=""5"">
    <page>
        
        <label x=""0"" y=""10"" width=""100"" height=""5"" brush=""black""/>
        <label x=""0"" y=""15"" width=""100"" height=""5"" brush=""black""/>
        <label x=""0"" y=""20"" width=""100"" height=""5"" brush=""black""/>
        <label x=""0"" y=""25"" width=""100"" height=""5"" brush=""black""/>
        <label x=""0"" y=""30"" width=""100"" height=""5"" brush=""black""/>
        <label x=""0"" y=""35"" width=""100"" height=""5"" brush=""black""/>
        <label x=""0"" y=""40"" width=""100"" height=""5"" brush=""black""/>
        <label x=""0"" y=""45"" width=""100"" height=""5"" brush=""black""/>
        <label x=""0"" y=""50"" width=""100"" height=""5"" brush=""black""/>
        <label x=""0"" y=""55"" width=""100"" height=""5"" brush=""black""/>
    </page>
</report>", a1.repdata)

            Dim pageXml2 =
"<report kind=""A4"" padding=""5"">
    <page>{foreach txt in ['あ', 'い', 'う']}
        {for i = base + 0 to base + 2}
        <label x=""0"" y=""#{i * 5}"" width=""100"" height=""5"" brush=""black"" text=""#{txt}""/>{/for}{/for}
    </page>
</report>"
            Dim a2 = ReadReportsInformation(pageXml2, New With {.base = 2})
            Assert.Equal("<report kind=""A4"" padding=""5"">
    <page>
        
        <label x=""0"" y=""10"" width=""100"" height=""5"" brush=""black"" text=""あ""/>
        <label x=""0"" y=""15"" width=""100"" height=""5"" brush=""black"" text=""あ""/>
        <label x=""0"" y=""20"" width=""100"" height=""5"" brush=""black"" text=""あ""/>
        
        <label x=""0"" y=""10"" width=""100"" height=""5"" brush=""black"" text=""い""/>
        <label x=""0"" y=""15"" width=""100"" height=""5"" brush=""black"" text=""い""/>
        <label x=""0"" y=""20"" width=""100"" height=""5"" brush=""black"" text=""い""/>
        
        <label x=""0"" y=""10"" width=""100"" height=""5"" brush=""black"" text=""う""/>
        <label x=""0"" y=""15"" width=""100"" height=""5"" brush=""black"" text=""う""/>
        <label x=""0"" y=""20"" width=""100"" height=""5"" brush=""black"" text=""う""/>
    </page>
</report>", a2.repdata)
        End Using

        Dim pageXml3 =
"<report kind=""A4"" padding=""5"">
    <page>{for i = base + 0 to base + 2}{foreach txt in texts}
        <label x=""0"" y=""#{i * 5}"" width=""100"" height=""5"" brush=""black"" text=""#{txt}""/>{/for}{/for}
    </page>
</report>"
        Dim a3 = ReadReportsInformation(pageXml3, New With {.base = 2, .texts = {"AA", "BB", "CC"}})
        Assert.Equal("<report kind=""A4"" padding=""5"">
    <page>
        <label x=""0"" y=""10"" width=""100"" height=""5"" brush=""black"" text=""AA""/>
        <label x=""0"" y=""10"" width=""100"" height=""5"" brush=""black"" text=""BB""/>
        <label x=""0"" y=""10"" width=""100"" height=""5"" brush=""black"" text=""CC""/>
        <label x=""0"" y=""15"" width=""100"" height=""5"" brush=""black"" text=""AA""/>
        <label x=""0"" y=""15"" width=""100"" height=""5"" brush=""black"" text=""BB""/>
        <label x=""0"" y=""15"" width=""100"" height=""5"" brush=""black"" text=""CC""/>
        <label x=""0"" y=""20"" width=""100"" height=""5"" brush=""black"" text=""AA""/>
        <label x=""0"" y=""20"" width=""100"" height=""5"" brush=""black"" text=""BB""/>
        <label x=""0"" y=""20"" width=""100"" height=""5"" brush=""black"" text=""CC""/>
    </page>
</report>", a3.repdata)
    End Sub

    '    <Fact>
    '    Public Sub SelectTest()
    '        Dim pageXml1 =
    '"<report kind=""A4"" padding=""5"">
    '    <page>
    '        <label x=""10"" y=""20"" width=""30"" height=""6"" brush=""black""/>
    '    </page>
    '</report>"
    '        Dim a1 = ReadReportsInformation(pageXml1)
    '        a1.CalcLocation()
    '    End Sub

    '    <Fact>
    '    Public Sub LabelTest()
    '        Dim pageXml1 =
    '"<report kind=""A4"" padding=""5"">
    '    <page>
    '        <label x=""10"" y=""20"" width=""30"" height=""6"" brush=""black""/>
    '    </page>
    '</report>"
    '        Dim a1 = ReadReportsInformation(pageXml1)
    '        a1.CalcLocation()
    '    End Sub

End Class
