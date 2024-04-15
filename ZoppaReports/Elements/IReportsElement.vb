Option Strict On
Option Explicit On

Imports System.Drawing
Imports ZoppaReports.Resources
Imports ZoppaReports.Smalls

''' <summary>要素インターフェース。</summary>
Public Interface IReportsElement
    Inherits IDisposable

    ''' <summary>親要素を取得します。</summary>
    ''' <returns>親要素。</returns>
    ReadOnly Property Parent As IReportsElement

    ''' <summary>子要素リストを取得します。</summary>
    ''' <returns>子要素リスト。</returns>
    ReadOnly Property Children As IList(Of IReportsElement)

    ''' <summary>リリースリストを取得します。</summary>
    ''' <returns>リリースリスト。</returns>
    ReadOnly Property Resources As IList(Of IReportsResources)

    ''' <summary>マージンを取得します。</summary>
    ReadOnly Property Margin As ReportsMargin

    ''' <summary>パディングを取得します。</summary>
    ReadOnly Property Padding As ReportsPadding

    ''' <summary>クリップする、しないを取得します。</summary>
    ReadOnly Property ClipToBounds As Boolean

    ''' <summary>リソースを取得します。</summary>
    ''' <param name="name">リソース名。</param>
    ''' <returns>リソース。</returns>
    Function GetResource(Of T As {IReportsResources})(name As String) As T

    ''' <summary>基本スタイルを取得します。</summary>
    ''' <param name="typeBase">適用する型。</param>
    ''' <returns>スタイルリソース。</returns>
    Function GetBaseStyle(typeBase As Type) As StyleResource

    ''' <summary>子要素を追加します。</summary>
    ''' <param name="child">子要素。</param>
    Sub AddChildElement(child As IReportsElement)

    ''' <summary>リソースを設定します。</summary>
    ''' <param name="addRes">リソース。</param>
    Sub AddResources(addRes As IReportsResources)

    ''' <summary>印刷を実行します。</summary>
    ''' <param name="graphics">グラフィックオブジェクト。</param>
    ''' <param name="clipBounds">クリップ領域。</param>
    Sub Draw(graphics As Graphics, clipBounds As Rectangle)

End Interface
