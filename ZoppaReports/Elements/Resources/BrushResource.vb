﻿Option Strict On
Option Explicit On

Imports System.Drawing
Imports ZoppaReports.Exceptions

Namespace Resources

    ''' <summary>ブラシリソース。</summary>
    Public NotInheritable Class BrushResource
        Implements IReportsResources

        ' リソース名
        Private ReadOnly mName As String

        ' ブラシ
        Private mBrush As Brush = Brushes.Black

        ' 前景色
        Private mColor As Color = Color.Black

        ' 背景色
        Private mBkColor As Color = Color.Transparent

        ' ハッチスタイル
        Private mHatchStyle As Drawing2D.HatchStyle? = Nothing

#Region "properties"

        ''' <summary>リソース名を取得します。</summary>
        ''' <returns>リソース名、</returns>
        Public ReadOnly Property Name As String Implements IReportsResources.Name
            Get
                Return Me.mName
            End Get
        End Property

        ''' <summary>リソースを取得します。</summary>
        ''' <returns>リソース。</returns>
        Private Function Contents(Of T As {Class})() As T Implements IReportsResources.Contents
            If Me.mBrush Is Nothing Then
                Me.mBrush = If(Me.mHatchStyle Is Nothing,
                    CType(New SolidBrush(Me.mColor), Brush),
                    New Drawing2D.HatchBrush(Me.mHatchStyle.Value, Me.mColor, Me.mBkColor)
                )
            End If

            Dim res = TryCast(Me.mBrush, T)
            If res IsNot Nothing Then
                Return res
            Else
                Throw New InvalidCastException($"リソースは{GetType(T).Name}ではありません")
            End If
        End Function

        ''' <summary>前景色を取得します。</summary>
        ''' <returns>前景色。</returns>
        Public ReadOnly Property ForeColor As Color
            Get
                Return Me.mColor
            End Get
        End Property

        ''' <summary>背景色を取得します。</summary>
        ''' <returns>背景色。</returns>
        Public ReadOnly Property BackColor As Color
            Get
                Return Me.mBkColor
            End Get
        End Property

        ''' <summary>ハッチスタイルを取得します。</summary>
        ''' <returns>ハッチスタイル。</returns>
        Public ReadOnly Property HatchStyle As Drawing2D.HatchStyle?
            Get
                Return Me.mHatchStyle
            End Get
        End Property

#End Region

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="nm">リソース名。</param>
        Public Sub New(nm As String)
            Me.mName = nm
        End Sub

        ''' <summary>リソースを解放します。</summary>
        Public Sub Dispose() Implements IDisposable.Dispose
            Me.mBrush?.Dispose()
            Me.mBrush = Nothing
        End Sub

        ''' <summary>プロパティを設定します。</summary>
        ''' <param name="name">プロパティ名。</param>
        ''' <param name="value">プロパティ値。</param>
        ''' <returns>追加できたら真。</returns>
        Public Function SetProperty(name As String, value As Object) As Boolean Implements IReportsResources.SetProperty
            Me.mBrush?.Dispose()
            Me.mBrush = Nothing

            Select Case name.ToLower()
                Case NameOf(Me.ForeColor).ToLower(), "color"
                    Me.mColor = ConvertColor(value)
                    Return True

                Case NameOf(Me.BackColor).ToLower()
                    Me.mBkColor = ConvertColor(value)
                    Return True

                Case NameOf(Me.HatchStyle).ToLower()
                    Dim hBrh As Drawing2D.HatchStyle
                    If [Enum].TryParse(Of Drawing2D.HatchStyle)(value.ToString(), True, hBrh) Then
                        Me.mHatchStyle = hBrh
                        Return True
                    Else
                        Throw New ReportsAnalysisException($"{NameOf(Me.HatchStyle)}プロパティに{value.GetType().Name}を設定できません")
                    End If

                Case Else
                    Return False
            End Select
        End Function

    End Class

End Namespace

