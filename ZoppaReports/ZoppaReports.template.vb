Option Strict On
Option Explicit On

Imports System.Runtime.CompilerServices
Imports System.Xml
Imports ZoppaReports.Exceptions

''' <summary>帳票モジュール。</summary>
Partial Public Module ZoppaReports

    ''' <summary>XMLノードの属性より情報を取得する(Double)</summary>
    ''' <param name="attr">属税。</param>
    ''' <returns>情報。</returns>
    <Extension()>
    Friend Function GetValueDouble(attr As XmlAttribute) As (has As Boolean, value As Double)
        If attr IsNot Nothing Then
            Try
                Return (True, Convert.ToDouble(attr.Value))
            Catch ex As Exception
                Throw New ReportsException($"属性値の変換に失敗しました。{attr.Value}", ex)
            End Try
        Else
            Return (False, Nothing)
        End If
    End Function

    ''' <summary>XMLノードの属性より情報を取得する。</summary>
    ''' <typeparam name="T">取得した情報の型。</typeparam>
    ''' <param name="attr">属税。</param>
    ''' <returns>情報。</returns>
    <Extension()>
    Friend Function GetValue(Of T)(attr As XmlAttribute) As (has As Boolean, value As T)
        If attr IsNot Nothing Then
            Try
                Return (True, CTypeDynamic(Of T)(attr.Value))
            Catch ex As Exception
                Throw New ReportsException($"属性値の変換に失敗しました。{attr.Value}", ex)
            End Try
        Else
            Return (False, Nothing)
        End If
    End Function

    'Friend Sub ReadElements(parent As IReportsElements, nodes As Xml.XmlNodeList)
    '    For Each nd As Xml.XmlNode In nodes
    '        Select Case nd.Name
    '            Case "label"
    '        End Select
    '    Next
    'End Sub

End Module