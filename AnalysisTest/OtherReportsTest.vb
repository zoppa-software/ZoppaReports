Imports Xunit
Imports ZoppaReports.Designer

Public Class OtherReportsTest

    <Fact>
    Public Sub PageOrientationTest()
        Dim ai = CType(ReportsOrientation.Landscape, String)
        Assert.Equal("landscape", ai)

        Dim ai2 = CType(ReportsOrientation.Portrait, String)
        Assert.Equal("portrait", ai2)

        Dim ai3 = CType("landscape", ReportsOrientation)
        Assert.Equal(ReportsOrientation.Landscape, ai3)

        Dim ai4 = CType("portrait", ReportsOrientation)
        Assert.Equal(ReportsOrientation.Portrait, ai4)

        Dim a1 = (ReportsOrientation.Landscape = ReportsOrientation.Landscape)
        Assert.True(a1)

        Dim a11 = (ReportsOrientation.Landscape = Nothing)
        Assert.False(a11)

        Dim a12 = (Nothing = ReportsOrientation.Landscape)
        Assert.False(a12)

        Dim a2 = (ReportsOrientation.Landscape <> ReportsOrientation.Portrait)
        Assert.True(a2)

        Dim a21 = ReportsOrientation.Landscape.Equals(ReportsOrientation.Landscape)
        Assert.True(a21)

        Dim hc = ReportsOrientation.Landscape.GetHashCode()

        Assert.Throws(Of InvalidCastException)(Function() CType("xxx", ReportsOrientation))
    End Sub

    <Fact>
    Public Sub ReportsSizeTest()

    End Sub

End Class
