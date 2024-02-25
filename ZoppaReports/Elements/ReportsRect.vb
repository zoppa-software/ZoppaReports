Option Strict On
Option Explicit On

Imports ZoppaReports.Designer

Namespace Elements

    Public Structure ReportsRect



        Private mIsCalced As Boolean

        Private mLeft As Double, mTop As Double, mWidth As Double, mHeight As Double

        Public ReadOnly Property IsCalced As Boolean
            Get
                Return Me.mIsCalced
            End Get
        End Property

        Public ReadOnly Property Left As Double
            Get
                Return Me.mLeft
            End Get
        End Property

        Public ReadOnly Property Top As Double
            Get
                Return Me.mTop
            End Get
        End Property

        Public ReadOnly Property Width As Double
            Get
                Return Me.mWidth
            End Get
        End Property

        Public ReadOnly Property Height As Double
            Get
                Return Me.mHeight
            End Get
        End Property

        Public Sub Reset()
            Me.mIsCalced = False
            Me.mLeft = 0
            Me.mTop = 0
            Me.mWidth = 0
            Me.mHeight = 0
        End Sub

        Public Sub Calc(parent As IReportsElements)

        End Sub

    End Structure

End Namespace
