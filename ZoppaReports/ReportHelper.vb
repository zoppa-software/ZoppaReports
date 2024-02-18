Option Strict On
Option Explicit On
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Documents

Public Module ReportHelper

    Public Async Function ReadView(data As String) As Task(Of FixedDocument)
        Return Await ReadView(New IO.StringReader(data))
    End Function

    Public Async Function ReadView(filePath As IO.FileInfo, Optional enc As Text.Encoding = Nothing) As Task(Of FixedDocument)
        Using sr = New IO.StreamReader(filePath.FullName, If(enc, System.Text.Encoding.Default))
            Return Await ReadView(sr)
        End Using
    End Function

    Public Async Function ReadView(reader As IO.TextReader) As Task(Of FixedDocument)
        Dim data = Await reader.ReadToEndAsync()

        Dim doc As New FixedDocument()

        For i As Integer = 0 To 2
            Dim page = New FixedPage()
            Dim canvas = New Canvas()
            Dim tb = New TextBlock()
            tb.Text = String.Format("{0}ページ", i + 1)
            tb.FontSize = 24
            Canvas.SetTop(tb, 100)
            Canvas.SetLeft(tb, 100)
            canvas.Children.Add(tb)
            page.Children.Add(canvas)

            Dim pc = New PageContent()
            pc.Child = page
            doc.Pages.Add(pc)
        Next

        '    FixedPage page = New FixedPage();
        'Canvas canvas = New Canvas();
        'TextBlock tb = New TextBlock();
        'tb.Text = String.Format("{0}ページ", i + 1);
        'tb.FontSize = 24;
        'Canvas.SetTop(tb, 100);
        'Canvas.SetLeft(tb, 100);
        'canvas.Children.Add(tb);
        'page.Children.Add(canvas);

        'PageContent pc = New PageContent();
        'pc.Child = page;
        'doc.Pages.Add(pc);

        Return doc
    End Function

End Module
