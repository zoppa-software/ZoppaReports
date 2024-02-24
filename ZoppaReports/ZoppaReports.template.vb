Option Strict On
Option Explicit On

Imports Microsoft.Extensions.Logging
Imports Microsoft.Extensions.DependencyInjection
Imports System.Runtime.CompilerServices
Imports System.Windows.Documents
Imports System.Drawing
Imports System.Drawing.Printing
Imports System.Xml
Imports System.Text
Imports System.Diagnostics.Eventing
Imports System.Drawing.Drawing2D
Imports ZoppaReports.Exceptions

''' <summary>帳票モジュール。</summary>
Partial Public Module ZoppaReports

    ''' <summary>文字列の置き換え式の置き換えを行う。</summary>
    ''' <param name="source">文字列。</param>
    ''' <param name="parameter">環境値。</param>
    ''' <returns>置き換え後の文字列。</returns>
    Private Function ReplaceExpressions(source As String, parameter As Object) As String
        Dim res As New StringBuilder()
        Dim buffer As New StringBuilder()

        Dim pos As Integer = 0
        Dim replase As Boolean = False
        Do While pos < source.Length
            Dim c = source(pos)

            Select Case c
                Case "}"c
                    If replase Then
                        res.Append(CTypeDynamic(Of String)(buffer.ToString().Executes(parameter).Contents))
                        buffer.Clear()
                        replase = False
                        pos += 1
                    Else
                        buffer.Append(c)
                        pos += 1
                    End If

                Case "#"c
                    Dim nc = If(pos + 1 < source.Length, source(pos + 1), Nothing)
                    If nc = "{"c Then
                        res.Append(buffer)
                        buffer.Clear()
                        replase = True
                        pos += 2
                    Else
                        buffer.Append(c)
                        pos += 1
                    End If

                Case "\"c
                    Dim nc = If(pos + 1 < source.Length, source(pos + 1), Nothing)
                    If nc = "{"c OrElse nc = "}"c Then
                        buffer.Append(nc)
                        pos += 2
                    Else
                        buffer.Append(c)
                        pos += 1
                    End If

                Case Else
                    buffer.Append(c)
                    pos += 1
            End Select
        Loop

        If Not replase Then
            res.Append(buffer)
        Else
            Throw New ReportsException("置き換えトークンが閉じられていません")
        End If

        Return res.ToString()
    End Function

    ''' <summary>XMLノードの属性より情報を取得する(文字列)</summary>
    ''' <param name="attr">属税。</param>
    ''' <param name="parameter">環境パラメータ。</param>
    ''' <returns>情報。</returns>
    <Extension()>
    Friend Function GetValueString(attr As XmlAttribute, parameter As Object) As (has As Boolean, value As String)
        If attr IsNot Nothing Then
            Dim value = ReplaceExpressions(attr.Value, parameter)
            Return (True, value)
        Else
            Return (False, Nothing)
        End If
    End Function

    ''' <summary>XMLノードの属性より情報を取得する(Integer)</summary>
    ''' <param name="attr">属税。</param>
    ''' <param name="parameter">環境パラメータ。</param>
    ''' <returns>情報。</returns>
    <Extension()>
    Friend Function GetValueInteger(attr As XmlAttribute, parameter As Object) As (has As Boolean, value As Integer)
        If attr IsNot Nothing Then
            Dim value = ReplaceExpressions(attr.Value, parameter)
            Return (True, Convert.ToInt32(value))
        Else
            Return (False, Nothing)
        End If
    End Function

    ''' <summary>XMLノードの属性より情報を取得する(Double)</summary>
    ''' <param name="attr">属税。</param>
    ''' <param name="parameter">環境パラメータ。</param>
    ''' <returns>情報。</returns>
    <Extension()>
    Friend Function GetValueDouble(attr As XmlAttribute, parameter As Object) As (has As Boolean, value As Double)
        If attr IsNot Nothing Then
            Dim value = ReplaceExpressions(attr.Value, parameter)
            Return (True, Convert.ToDouble(value))
        Else
            Return (False, Nothing)
        End If
    End Function

    ''' <summary>XMLノードの属性より情報を取得する。</summary>
    ''' <typeparam name="T">取得した情報の型。</typeparam>
    ''' <param name="attr">属税。</param>
    ''' <param name="parameter">環境パラメータ。</param>
    ''' <returns>情報。</returns>
    <Extension()>
    Friend Function GetValue(Of T)(attr As XmlAttribute, parameter As Object) As (has As Boolean, value As T)
        If attr IsNot Nothing Then
            Dim value = ReplaceExpressions(attr.Value, parameter)
            Return (True, CTypeDynamic(Of T)(value))
        Else
            Return (False, Nothing)
        End If
    End Function

End Module