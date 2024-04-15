Imports System.Drawing
Imports Xunit
Imports ZoppaReports
Imports ZoppaReports.Draw

Public Class DrawTest

    <Fact>
    Public Sub LabelTest()
        Dim doc As New Xml.XmlDocument()
        doc.LoadXml("<label text='123' font='@{lb font}' >
    <resources>
        <font name='lb font' family='Meiryo UI' emsize='9.0' style='Italic' />
    </resources>
</label>")
        Dim ln = GetElementFromXml(Of LabelElement)(doc.FirstChild, Nothing)
        'Assert.Equal("def font", fn.Name)
        'Assert.Equal(9.0F, fn.EmSize)
        'Assert.Equal(FontStyle.Italic, fn.Style)
    End Sub

End Class
