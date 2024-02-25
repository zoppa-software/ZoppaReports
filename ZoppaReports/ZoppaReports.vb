Option Strict On
Option Explicit On

Imports Microsoft.Extensions.Logging
Imports Microsoft.Extensions.DependencyInjection
Imports System.Runtime.CompilerServices
Imports System.Windows.Documents
Imports System.Drawing
Imports System.Drawing.Printing
Imports ZoppaReports.Designer
Imports ZoppaReports.Lexical
Imports ZoppaReports.Parser

''' <summary>帳票モジュール。</summary>
Public Module ZoppaReports

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="printer"></param>
    ''' <param name="filePath"></param>
    ''' <param name="enc"></param>
    ''' <param name="parameter"></param>
    ''' <returns></returns>
    <Extension()>
    Public Async Function DrawReport(printer As PrintDocument,
                                     filePath As IO.FileInfo,
                                     Optional enc As System.Text.Encoding = Nothing,
                                     Optional parameter As Object = Nothing) As Task(Of ReportsInformation)
        Using sr = New IO.StreamReader(filePath.FullName, If(enc, System.Text.Encoding.Default))
            Return DrawReport(printer, Await sr.ReadToEndAsync())
        End Using
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="printer"></param>
    ''' <param name="data"></param>
    ''' <param name="parameter"></param>
    ''' <returns></returns>
    <Extension()>
    Public Function DrawReport(printer As PrintDocument,
                               data As String,
                               Optional parameter As Object = Nothing) As ReportsInformation
        Dim res = ReadReportsInformation(data, parameter)

        Dim rsiz = res.info.Size
        If rsiz.Kind = PaperKind.Custom Then
            printer.DefaultPageSettings.PaperSize = New PaperSize(
                rsiz.Kind.ToString(),
                CInt(Math.Round(rsiz.WidthInMM.MmToInch)),
                CInt(Math.Round(rsiz.HeightInMM.MmToInch))
            )
        Else
            For Each ps As PaperSize In printer.PrinterSettings.PaperSizes
                If ps.Kind = rsiz.Kind Then
                    printer.DefaultPageSettings.PaperSize = ps
                    Exit For
                End If
            Next
        End If

        printer.DefaultPageSettings.Landscape = (rsiz.Orientation = ReportsOrientation.Landscape)

        Dim handler As New PrintHandler(res.info)
        Try
            AddHandler printer.PrintPage, AddressOf handler.Draw
            printer.Print()
        Catch ex As Exception
        Finally
            AddHandler printer.PrintPage, AddressOf handler.Draw
        End Try

        Return res.info
    End Function

    Private NotInheritable Class PrintHandler

        Private ReadOnly mInfo As ReportsInformation

        Sub New(info As ReportsInformation)
            Me.mInfo = info
        End Sub

        Public Sub Draw(sender As Object, e As Printing.PrintPageEventArgs)
            Me.mInfo.Draw(e.Graphics)
        End Sub

    End Class

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="data"></param>
    ''' <param name="parameter"></param>
    ''' <returns></returns>
    Public Function ReadReportsInformation(data As String, Optional parameter As Object = Nothing) As (info As ReportsInformation, repdata As String)
        Dim env As New Environments(parameter)

        Using scope = _logger.Value?.BeginScope(NameOf(ReadReportsInformation))
            ' テンプレートをトークンに分割
            Dim splited = LexicalAnalysis.SplitRepToken(data)

            ' テンプレートを置き換え
            data = ParserAnalysis.Replase(data, splited, env)

            ' テンプレートを解析
            Return (ReportsInformation.Load(data), data)
        End Using
    End Function

End Module
