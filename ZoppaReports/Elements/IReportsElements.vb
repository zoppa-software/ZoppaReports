Option Strict On
Option Explicit On
Imports ZoppaReports.Designer

Namespace Elements

    ''' <summary>帳票要素。</summary>
    Public Interface IReportsElements

        ''' <summary>パディングを設定、取得します。</summary>
        ''' <returns>マージン情報。</returns>
        Property Padding As ReportsMargin

        ''' <summary>マージンを設定、取得します。</summary>
        ''' <returns>マージン情報。</returns>
        Property Margin As ReportsMargin

    End Interface

End Namespace
