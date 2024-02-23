Option Strict On
Option Explicit On

Imports Microsoft.Extensions.Logging
Imports Microsoft.Extensions.DependencyInjection
Imports System.Runtime.CompilerServices
Imports System.Windows.Documents
Imports System.Drawing
Imports System.Drawing.Printing
Imports System.Xml

''' <summary>帳票モジュール。</summary>
Partial Public Module ZoppaReports


    <Extension()>
    Public Function GetValue(Of T)(attr As XmlAttribute, parameter As Object) As (has As Boolean, value As T)
        If attr IsNot Nothing Then
            Dim value = attr.Value
            If value.Length > 3 AndAlso value.StartsWith("#{") AndAlso value.EndsWith("}") Then
                Dim o = value.Substring(2, value.Length - 3).Executes(parameter).Contents
                Return (True, CTypeDynamic(Of T)(o))
            ElseIf GetType(T) = GetType(Integer) Then
                Return (True, CType(CObj(Convert.ToInt32(value)), T))
            ElseIf GetType(T) = GetType(Double) Then
                Return (True, CType(CObj(Convert.ToDouble(value)), T))
            Else
                Return (True, CTypeDynamic(Of T)(value))
            End If
        Else
            Return (False, Nothing)
        End If
    End Function

End Module