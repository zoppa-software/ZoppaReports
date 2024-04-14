Option Strict On
Option Explicit On

Imports System.Drawing.Printing
Imports ZoppaReports

Public Class Form1

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles RunButton.Click
        Dim template = "" &
"<report size='B4' orientation='Landscape' >
    <resources>
        <font name='lb font' family='Meiryo UI' emsize='9.0' style='Italic' />
        <style base='label'>
            <set name='font' value='@{lb font}' />
        </style>
    </resources>
    <label text='123 abc 日本語' />
</report>"
        Me.PrintDoc.DrawReport(template)
    End Sub

End Class
