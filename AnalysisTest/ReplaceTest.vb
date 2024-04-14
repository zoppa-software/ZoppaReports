Imports System.Drawing
Imports System.Drawing.Printing
Imports System.Printing
Imports Microsoft.Extensions.Logging
Imports Xunit
Imports ZoppaReports
Imports ZoppaReports.Exceptions

Public Class ReplaceTest

    '    <Fact>
    '    Public Sub PageTest()
    '        Using logFacyory = LoggerFactory.Create(
    '                Sub(config)
    '                    config.SetMinimumLevel(LogLevel.Trace)
    '                    config.AddConsole()
    '                End Sub)
    '            SetZoppaReportsLogFactory(logFacyory)
    '            Dim pageXml =
    '"<report kind=""A4"">
    '            <resources>
    '                <font name=""MS Gothic"" size=""12"" />
    '            </resources>
    '            <page>
    '            </page>
    '        </report>
    '        "
    '            Dim a1 = ReadReportsInformation(pageXml)
    '            Assert.Equal(PaperKind.A4, a1.info.Size.Kind)

    '            Dim pageXml2 =
    '"<report kind=""#{kind}"">
    '            <page>
    '            </page>
    '        </report>
    '        "
    '            Dim a2 = ReadReportsInformation(pageXml2, New With {.kind = ReportsSize.B4})
    '            Assert.Equal(PaperKind.B4, a2.info.Size.Kind)

    '            Dim pageXml3 =
    '"<report width=""100"" height=""200"" orientation=""portrait"">
    '            <page>
    '            </page>
    '        </report>
    '        "
    '            Dim a3 = ReadReportsInformation(pageXml3)
    '            Assert.Equal(PaperKind.Custom, a3.info.Size.Kind)
    '            Assert.Equal(100, a3.info.Size.WidthInMM)
    '            Assert.Equal(200, a3.info.Size.HeightInMM)
    '            Assert.Equal(ReportsOrientation.Portrait, a3.info.Size.Orientation)

    '            Dim pageXml4 =
    '"<report kind=""custom,140,200,portrait"">
    '            <page>
    '            </page>
    '        </report>
    '        "
    '            Dim a4 = ReadReportsInformation(pageXml4)
    '            Assert.Equal(PaperKind.Custom, a4.info.Size.Kind)
    '            Assert.Equal(140, a4.info.Size.WidthInMM)
    '            Assert.Equal(200, a4.info.Size.HeightInMM)
    '            Assert.Equal(ReportsOrientation.Portrait, a4.info.Size.Orientation)

    '            Dim pageXml5 =
    '"<report width=""#{wid + 10}"" height=""#{hid - 10}"" orientation=""#{ori}"">
    '            <page>
    '            </page>
    '        </report>
    '        "
    '            Dim a5 = ReadReportsInformation(pageXml5, New With {.wid = 90, .hid = 110, .ori = ReportsOrientation.Landscape})
    '            Assert.Equal(PaperKind.Custom, a5.info.Size.Kind)
    '            Assert.Equal(100, a5.info.Size.WidthInMM)
    '            Assert.Equal(100, a5.info.Size.HeightInMM)
    '            Assert.Equal(ReportsOrientation.Landscape, a5.info.Size.Orientation)
    '        End Using
    '    End Sub

    '<Fact>
    'Public Sub PageErrorTest()
    '    '        Using logFacyory = LoggerFactory.Create(
    '    '                Sub(config)
    '    '                    config.SetMinimumLevel(LogLevel.Trace)
    '    '                    config.AddConsole()
    '    '                End Sub)
    '    '            SetZoppaReportsLogFactory(logFacyory)
    '    '            Dim pageXml =
    '    '"<reporterr kind=""A4"">
    '    '    <resources>
    '    '        <font name=""MS Gothic"" size=""12"" />
    '    '    </resources>
    '    '    <page>
    '    '    </page>
    '    '</reporterr>
    '    '"
    '    '            ' report要素未定義
    '    '            Assert.Throws(Of ReportsException)(
    '    '                Sub()
    '    '                    ReadReportsInformation(pageXml)
    '    '                End Sub
    '    '            )
    '    '        End Using
    'End Sub

    '<Fact>
    'Public Sub MarginTest()
    '    '        Using logFacyory = LoggerFactory.Create(
    '    '                Sub(config)
    '    '                    config.SetMinimumLevel(LogLevel.Trace)
    '    '                    config.AddConsole()
    '    '                End Sub)
    '    '            SetZoppaReportsLogFactory(logFacyory)
    '    '            Dim pageXml =
    '    '"<report kind=""A4"" padding=""3"">
    '    '    <page>
    '    '    </page>
    '    '</report>
    '    '"
    '    '            Dim a1 = ReadReportsInformation(pageXml)
    '    '            Assert.Equal(New ReportsMargin With {.Left = 3, .Top = 3, .Right = 3, .Bottom = 3}, a1.info.Padding)

    '    '            Dim pageXml1 =
    '    '    "<report kind=""A4"" padding=""3, 5"">
    '    '    <page>
    '    '    </page>
    '    '</report>
    '    '"
    '    '            Dim a2 = ReadReportsInformation(pageXml1)
    '    '            Assert.Equal(New ReportsMargin With {.Left = 3, .Top = 5, .Right = 3, .Bottom = 5}, a2.info.Padding)

    '    '            Dim pageXml2 =
    '    '    "<report kind=""A4"" padding=""1, 2, 3, 4"">
    '    '    <page>
    '    '    </page>
    '    '</report>
    '    '"
    '    '            Dim a3 = ReadReportsInformation(pageXml2)
    '    '            Assert.Equal(New ReportsMargin With {.Left = 1, .Top = 2, .Right = 3, .Bottom = 4}, a3.info.Padding)
    '    '        End Using
    'End Sub



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

    <Fact>
    Public Sub IfTest()
        Dim pageXml1 =
"<report kind=""A4"" padding=""5"">
    {if no < 10}
    <label x=""10"" y=""20"" width=""30"" height=""6"" brush=""black""/>
    {elseif no < 20}
    <label x=""40"" y=""50"" width=""60"" height=""12"" brush=""black""/>
    {else}
    <label x=""70"" y=""80"" width=""90"" height=""18"" brush=""black""/>
    {end if}
</report>"
        Dim a1 = ReplaceTemplateXml(pageXml1, New Environments(New With {.no = 5}))
        Assert.Equal("<report kind=""A4"" padding=""5"">
    
    <label x=""10"" y=""20"" width=""30"" height=""6"" brush=""black""/>
    
</report>", a1)

        Dim a2 = ReplaceTemplateXml(pageXml1, New Environments(New With {.no = 19}))
        Assert.Equal("<report kind=""A4"" padding=""5"">
    
    <label x=""40"" y=""50"" width=""60"" height=""12"" brush=""black""/>
    
</report>", a2)

        Dim a3 = ReplaceTemplateXml(pageXml1, New Environments(New With {.no = 30}))
        Assert.Equal("<report kind=""A4"" padding=""5"">
    
    <label x=""70"" y=""80"" width=""90"" height=""18"" brush=""black""/>
    
</report>", a3)
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
    {for i = base + 0 to base + 9}
    <label x=""0"" y=""{i * 5}"" width=""100"" height=""5"" brush=""black""/>{/for}
</report>"
            Dim a1 = ReplaceTemplateXml(pageXml1, New Environments(New With {.base = 2}))
            Assert.Equal("<report kind=""A4"" padding=""5"">
    
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
</report>", a1)

            Dim pageXml2 =
"<report kind=""A4"" padding=""5"">
    {foreach txt in ['あ', 'い', 'う']}{for i = base + 0 to base + 2}
    <label x=""0"" y=""{i * 5}"" width=""100"" height=""5"" brush=""black"" text=""{txt}""/>{/for}{/for}
</report>"
            Dim a2 = ReplaceTemplateXml(pageXml2, New Environments(New With {.base = 2}))
            Assert.Equal("<report kind=""A4"" padding=""5"">
    
    <label x=""0"" y=""10"" width=""100"" height=""5"" brush=""black"" text=""あ""/>
    <label x=""0"" y=""15"" width=""100"" height=""5"" brush=""black"" text=""あ""/>
    <label x=""0"" y=""20"" width=""100"" height=""5"" brush=""black"" text=""あ""/>
    <label x=""0"" y=""10"" width=""100"" height=""5"" brush=""black"" text=""い""/>
    <label x=""0"" y=""15"" width=""100"" height=""5"" brush=""black"" text=""い""/>
    <label x=""0"" y=""20"" width=""100"" height=""5"" brush=""black"" text=""い""/>
    <label x=""0"" y=""10"" width=""100"" height=""5"" brush=""black"" text=""う""/>
    <label x=""0"" y=""15"" width=""100"" height=""5"" brush=""black"" text=""う""/>
    <label x=""0"" y=""20"" width=""100"" height=""5"" brush=""black"" text=""う""/>
</report>", a2)


            Dim pageXml3 =
"<report kind=""A4"" padding=""5"">
    {for i = base + 0 to base + 2}{foreach txt in texts}
    <label x=""0"" y=""{i * 5}"" width=""100"" height=""5"" brush=""black"" text=""{txt}""/>{/for}{/for}
</report>"
            Dim a3 = ReplaceTemplateXml(pageXml3, New Environments(New With {.base = 2, .texts = {"AA", "BB", "CC"}}))
            Assert.Equal("<report kind=""A4"" padding=""5"">
    
    <label x=""0"" y=""10"" width=""100"" height=""5"" brush=""black"" text=""AA""/>
    <label x=""0"" y=""10"" width=""100"" height=""5"" brush=""black"" text=""BB""/>
    <label x=""0"" y=""10"" width=""100"" height=""5"" brush=""black"" text=""CC""/>
    <label x=""0"" y=""15"" width=""100"" height=""5"" brush=""black"" text=""AA""/>
    <label x=""0"" y=""15"" width=""100"" height=""5"" brush=""black"" text=""BB""/>
    <label x=""0"" y=""15"" width=""100"" height=""5"" brush=""black"" text=""CC""/>
    <label x=""0"" y=""20"" width=""100"" height=""5"" brush=""black"" text=""AA""/>
    <label x=""0"" y=""20"" width=""100"" height=""5"" brush=""black"" text=""BB""/>
    <label x=""0"" y=""20"" width=""100"" height=""5"" brush=""black"" text=""CC""/>
</report>", a3)
        End Using
    End Sub

    <Fact>
    Public Sub SelectTest()
        Dim pageXml1 =
"<report kind=""A4"" padding=""5"">
    {select lbl}
    {case 'A'}
    <label x=""10"" y=""20"" width=""30"" height=""6"" brush=""black""/>
    {case 'B'}
    <label x=""40"" y=""50"" width=""60"" height=""12"" brush=""black""/>
    {else}
    <label x=""70"" y=""80"" width=""90"" height=""18"" brush=""black""/>
    {/select}
</report>"
        Dim a1 = ReplaceTemplateXml(pageXml1, New Environments(New With {.lbl = "A"}))
        Assert.Equal("<report kind=""A4"" padding=""5"">
    
    <label x=""10"" y=""20"" width=""30"" height=""6"" brush=""black""/>
    
</report>", a1)

        Dim a2 = ReplaceTemplateXml(pageXml1, New Environments(New With {.lbl = "B"}))
        Assert.Equal("<report kind=""A4"" padding=""5"">
    
    <label x=""40"" y=""50"" width=""60"" height=""12"" brush=""black""/>
    
</report>", a2)

        Dim a3 = ReplaceTemplateXml(pageXml1, New Environments(New With {.lbl = "C"}))
        Assert.Equal("<report kind=""A4"" padding=""5"">
    
    <label x=""70"" y=""80"" width=""90"" height=""18"" brush=""black""/>
    
</report>", a3)
    End Sub

End Class
