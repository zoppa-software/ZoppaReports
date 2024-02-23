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

        Dim a2 = (ReportsOrientation.Landscape = ReportsOrientation.Portrait)
        Assert.False(a2)

        Assert.Throws(Of InvalidCastException)(Function() CType("xxx", ReportsOrientation))
    End Sub

End Class
