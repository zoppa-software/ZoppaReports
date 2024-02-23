Option Strict On
Option Explicit On

Imports Microsoft.Extensions.Logging
Imports Microsoft.Extensions.DependencyInjection
Imports System.Runtime.CompilerServices
Imports System.Windows.Documents
Imports System.Drawing
Imports System.Drawing.Printing
Imports ZoppaReports.Designer

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

        Dim rsiz = res.Size
        If rsiz.Kind = PaperKind.Custom Then
            printer.DefaultPageSettings.PaperSize = New PaperSize(
                rsiz.Kind.ToString(),
                CInt(Math.Round(rsiz.WidthInInc)),
                CInt(Math.Round(rsiz.HeightInInc))
            )
        Else
            For Each ps As PaperSize In printer.PrinterSettings.PaperSizes
                If ps.Kind = rsiz.Kind Then
                    printer.DefaultPageSettings.PaperSize = ps
                    Exit For
                End If
            Next
        End If

        printer.DefaultPageSettings.Landscape = (rsiz.Orientation = PageOrientation.Landscape)

        Dim handler As New PrintHandler(res)
        Try
            AddHandler printer.PrintPage, AddressOf handler.Draw
            printer.Print()
        Catch ex As Exception
        Finally
            AddHandler printer.PrintPage, AddressOf handler.Draw
        End Try

        Return res
    End Function

    Private NotInheritable Class PrintHandler

        Private ReadOnly mInfo As ReportsInformation

        Sub New(info As ReportsInformation)
            Me.mInfo = info
        End Sub

        Public Async Sub Draw(sender As Object, e As Printing.PrintPageEventArgs)
            Await Task.Run(
            Sub()
                Me.mInfo.Draw(e.Graphics)
            End Sub
        )
        End Sub

    End Class

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="data"></param>
    ''' <param name="parameter"></param>
    ''' <returns></returns>
    Public Function ReadReportsInformation(data As String, Optional parameter As Object = Nothing) As ReportsInformation
        ' Replase

        Return ReportsInformation.Load(data, parameter)
    End Function

End Module
