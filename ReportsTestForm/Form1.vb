Option Strict On
Option Explicit On

Imports System.Drawing.Printing
Imports ZoppaReports
Imports ZoppaReports.Designer

Public Class Form1

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles RunButton.Click
        Dim template =
"<report kind=""#{size}"" orientation=""landscape"">
    <resources>
        <font name=""MS Gothic"" size=""12"" />
    </resources>
    <page>
    </page>
</report>"
        Me.PrintDoc.DrawReport(template, New With {.size = ReportsSize.B4})
    End Sub

End Class
