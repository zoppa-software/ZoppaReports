Option Strict On
Option Explicit On

Imports ZoppaReports.Designer

Namespace Elements

    ''' <summary>帳票要素。</summary>
    Public Interface IReportsElements

        ''' <summary>子要素リストを取得します。</summary>
        ''' <returns>子要素リスト。</returns>
        ReadOnly Property Elements As List(Of IReportsElements)

        ''' <summary>パディングを設定、取得します。</summary>
        ''' <returns>マージン情報。</returns>
        Property Padding As ReportsMargin

        ''' <summary>マージンを設定、取得します。</summary>
        ''' <returns>マージン情報。</returns>
        Property Margin As ReportsMargin

        ''' <summary>左端位置を設定、取得します。</summary>
        ''' <returns>座標。</returns>
        Property LeftMM As Double

        ''' <summary>上端位置を設定、取得します。</summary>
        ''' <returns>座標。</returns>
        Property TopMM As Double

        ''' <summary>幅を設定、取得します。</summary>
        ''' <returns>座標。</returns>
        Property WidthMM As Double

        ''' <summary>高さを設定、取得します。</summary>
        ''' <returns>座標。</returns>
        Property HeightMM As Double

    End Interface

End Namespace
