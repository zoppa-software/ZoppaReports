Option Strict On
Option Explicit On

Imports ZoppaReports.Exceptions

Namespace Smalls

    ''' <summary>帳票の余白。</summary>
    Public Structure ReportsMargin

        ' 左
        Private mLeft As ReportsScalar

        ' 上
        Private mTop As ReportsScalar

        ' 右
        Private mRight As ReportsScalar

        ' 下
        Private mBottom As ReportsScalar

#Region "properties"

        ''' <summary>左余白を取得します。</summary>
        Public ReadOnly Property Left As ReportsScalar
            Get
                Return Me.mLeft
            End Get
        End Property

        ''' <summary>上余白を取得します。</summary>
        Public ReadOnly Property Top As ReportsScalar
            Get
                Return Me.mTop
            End Get
        End Property

        ''' <summary>右余白を取得します。</summary>
        Public ReadOnly Property Right As ReportsScalar
            Get
                Return Me.mRight
            End Get
        End Property

        ''' <summary>下余白を取得します。</summary>
        Public ReadOnly Property Bottom As ReportsScalar
            Get
                Return Me.mBottom
            End Get
        End Property

#End Region

        ''' <summary>文字列からPageOrientationに変換する</summary>
        ''' <param name="inp">入力文字列。</param>
        ''' <returns>帳票の向き。</returns>
        Public Shared Widening Operator CType(inp As String) As ReportsMargin
            Dim values = inp.Split(","c)
            Dim res As New ReportsMargin()

            Select Case values.Length
                Case 1
                    Dim v1 = CType(values(0), ReportsScalar)
                    res.mLeft = v1
                    res.mTop = v1
                    res.mRight = v1
                    res.mBottom = v1

                Case 2
                    Dim vh = CType(values(0), ReportsScalar)
                    Dim vv = CType(values(1), ReportsScalar)
                    res.mLeft = vh
                    res.mTop = vv
                    res.mRight = vh
                    res.mBottom = vv

                Case 4
                    res.mLeft = CType(values(0), ReportsScalar)
                    res.mTop = CType(values(1), ReportsScalar)
                    res.mRight = CType(values(2), ReportsScalar)
                    res.mBottom = CType(values(3), ReportsScalar)

                Case Else
                    Throw New ReportsException("余白の指定が不正です")
            End Select

            Return res
        End Operator

    End Structure

End Namespace
