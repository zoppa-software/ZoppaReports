Option Strict On
Option Explicit On

Imports System.Drawing
Imports System.Drawing.Printing
Imports System.Runtime.CompilerServices
Imports Microsoft.Extensions.Logging
Imports ZoppaReports.Exceptions
Imports ZoppaReports.Lexical
Imports ZoppaReports.Parser
Imports ZoppaReports.Smalls

''' <summary>帳票モジュール。</summary>
Public Module ZoppaReports

    ''' <summary>帳票を描画します。</summary>
    ''' <param name="printer">印刷ドキュメント。</param>
    ''' <param name="filePath">ファイルパス。</param>
    ''' <param name="enc">エンコード。</param>
    ''' <param name="parameter">パラメータ。</param>
    <Extension()>
    Public Async Sub DrawReport(printer As PrintDocument,
                                filePath As IO.FileInfo,
                                Optional enc As System.Text.Encoding = Nothing,
                                Optional parameter As Object = Nothing)
        Using sr = New IO.StreamReader(filePath.FullName, If(enc, System.Text.Encoding.Default))
            DrawReport(printer, Await sr.ReadToEndAsync(), parameter)
        End Using
    End Sub

    ''' <summary>帳票を描画します。</summary>
    ''' <param name="printer">印刷ドキュメント。</param>
    ''' <param name="template">テンプレート。</param>
    ''' <param name="parameter">パラメータ。</param>
    <Extension()>
    Public Sub DrawReport(printer As PrintDocument,
                          template As String,
                          Optional parameter As Object = Nothing)
        ' レポート情報を取得
        Dim analisysXml = CreateReportsStructFromXml(template, parameter)

        SettingPrintDocument(printer, analisysXml.info)

        Using scope = _logger.Value?.BeginScope(NameOf(ZoppaReports))
            If printer.PrintController Is Nothing Then
                printer.PrintController = New StandardPrintController()
            End If

            Dim handler As New PrintHandler(printer, analisysXml.info)
            AddHandler printer.PrintPage, AddressOf handler.Draw
            printer.Print()
        End Using
    End Sub

    ''' <summary>印刷ドキュメントの設定を行います。</summary>
    ''' <param name="printer">印刷ドキュメント。</param>
    ''' <param name="report">レポート情報。</param>
    Private Sub SettingPrintDocument(printer As PrintDocument, report As IReportsElement)
        Dim repDoc = TryCast(report, ReportsElement)

        ' 用紙設定
        If repDoc.PaperSize.Kind = PaperKind.Custom Then
            printer.DefaultPageSettings.PaperSize = New PaperSize(
                repDoc.PaperSize.Kind.ToString(),
                CInt(repDoc.PaperSize.Width.Inch * 100),
                CInt(repDoc.PaperSize.Height.Inch * 100)
            )
        Else
            For Each ps As PaperSize In printer.PrinterSettings.PaperSizes
                If ps.Kind = repDoc.PaperSize.Kind Then
                    printer.DefaultPageSettings.PaperSize = ps
                    Exit For
                End If
            Next
        End If

        ' 用紙向き
        printer.DefaultPageSettings.Landscape = (repDoc.PaperOrientation = ReportsOrientation.Landscape)
    End Sub

    ''' <summary>印刷イベントハンドラ。</summary>
    Private NotInheritable Class PrintHandler

        ' 印刷ドキュメント
        Private ReadOnly mPrinter As PrintDocument

        ' レポート情報
        Private ReadOnly mInfo As IReportsElement

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="printer">印刷ドキュメント。</param>
        ''' <param name="info">レポート情報。</param>
        Sub New(printer As PrintDocument, info As IReportsElement)
            Me.mPrinter = printer
            Me.mInfo = info
        End Sub

        ''' <summary>描画処理。</summary>
        ''' <param name="sender">送信元。</param>
        ''' <param name="e">イベント引数。</param>
        Public Sub Draw(sender As Object, e As Printing.PrintPageEventArgs)
            Using scope = _logger.Value?.BeginScope(NameOf(ZoppaReports))
                Try
                    Logger.Value?.LogInformation("draw start")
                    Me.mInfo.Draw(e.Graphics, Rectangle.Empty)
                    Logger.Value?.LogInformation("draw end")

                Catch ex As Exception
                    Logger.Value?.LogCritical("report drawing error : {Message}", ex.Message)
                Finally
                    RemoveHandler Me.mPrinter.PrintPage, AddressOf Me.Draw
                End Try
            End Using
        End Sub
    End Class

    ''' <summary>レポート要素をXMLから読み込みます。</summary>
    ''' <param name="templateXml">テンプレートXML。</param>
    ''' <param name="parameter">パラメータ。</param>
    ''' <returns>レポート要素。</returns>
    Public Function CreateReportsStructFromXml(templateXml As String,
                                               Optional parameter As Object = Nothing) As (info As IReportsElement, repdata As String)
        Dim env As New Environments(parameter)

        Using scope = _logger.Value?.BeginScope(NameOf(ZoppaReports))
            Logger.Value?.LogDebug("answer input template : {templateXml}", templateXml)

            ' テンプレートを置き換え
            Dim anserXml = ReplaceTemplateXml(templateXml, env)

            ' XML情報から読み込む
            Dim doc = New Xml.XmlDocument()
            doc.LoadXml(anserXml)

            ' XMLからレポート要素を取得する
            Dim root = doc.SelectSingleNode("report")
            If root IsNot Nothing Then
                Return (GetElementFromXml(root, Nothing, env), anserXml)
            Else
                Throw New ReportsException("report要素が設定されていません")
            End If
        End Using
    End Function

    ''' <summary>テンプレートを置き換えます。</summary>
    ''' <param name="templateXml">テンプレートXML。</param>
    ''' <param name="env">環境。</param>
    ''' <returns>置き換え後のXML。</returns>
    Public Function ReplaceTemplateXml(templateXml As String, env As Environments) As String
        ' テンプレートをトークンに分割
        Dim splited = LexicalAnalysis.SplitRepToken(templateXml)
        LoggingTokens(splited)

        ' テンプレートを置き換え
        Dim anserXml = ParserAnalysis.Replase(templateXml, splited, env)
        Logger.Value?.LogDebug("answer compile template : {anserXml}", anserXml)

        Return anserXml
    End Function

End Module
