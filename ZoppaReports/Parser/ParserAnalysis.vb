Option Strict On
Option Explicit On

Imports System.Text
Imports ZoppaReports.Exceptions
Imports ZoppaReports.Express
Imports ZoppaReports.Tokens

''' <summary>トークンを解析する。</summary>
Friend Module ParserAnalysis

    ''' <summary>式を解析して結果を取得します。</summary>
    ''' <param name="tokens">対象トークン。。</param>
    ''' <param name="parameter">パラメータ。</param>
    ''' <returns>解析結果。</returns>
    Friend Function Executes(tokens As List(Of TokenPosition), parameter As Environments) As (ans As IToken, expr As IExpression)
        ' 式木を作成
        Dim logicalParser As New LogicalParser()
        Dim compParser As New ComparisonParser()
        Dim addOrSubParser As New AddOrSubParser()
        Dim multiOrDivParser As New MultiOrDivParser()
        Dim refParser As New ReferenceParser()
        Dim facParser As New FactorParser()
        Dim parenParser As New ParenParser()

        ' 解析クラスを構成
        logicalParser.NextParser = compParser
        compParser.NextParser = addOrSubParser
        addOrSubParser.NextParser = multiOrDivParser
        multiOrDivParser.NextParser = refParser
        refParser.NextParser = facParser
        facParser.NextParser = parenParser
        parenParser.NextParser = logicalParser

        ' トークン解析
        Dim tknPtr = New TokenStream(tokens)
        Dim expr = logicalParser.Parser(tknPtr)

        ' 結果を取得する
        If Not tknPtr.HasNext Then
            Return (expr.Executes(parameter), expr)
        Else
            Throw New ReportsAnalysisException("未評価のトークンがあります")
        End If
    End Function

    ''' <summary>括弧内部式を取得します。</summary>
    ''' <param name="reader">入力トークンストリーム。</param>
    ''' <param name="nxtParser">次のパーサー。</param>
    ''' <returns>括弧内部式。</returns>
    Private Function CreateParenExpress(Of TLParen, TRParen)(reader As TokenStream, nxtParser As IParser) As ParenExpress
        Dim tmp As New List(Of TokenPosition)()
        Dim lv As Integer = 0
        Do While reader.HasNext
            Dim tkn = reader.Current
            reader.Move(1)

            Select Case tkn.TokenType
                Case GetType(TLParen)
                    tmp.Add(tkn)
                    lv += 1

                Case GetType(TRParen)
                    If lv > 0 Then
                        tmp.Add(tkn)
                        lv -= 1
                    Else
                        Exit Do
                    End If

                Case Else
                    tmp.Add(tkn)
            End Select
        Loop
        Return New ParenExpress(nxtParser.Parser(New TokenStream(tmp)))
    End Function

    ''' <summary>入力トークンストリーム。</summary>
    Friend NotInheritable Class TokenStream

        ''' <summary>シーク位置ポインタ。</summary>
        Private mPointer As Integer

        ''' <summary>入力トークン。</summary>
        Private ReadOnly mTokens As List(Of TokenPosition)

        ''' <summary>読み込みの終了していない文字があれば真を返す。</summary>
        Public ReadOnly Property HasNext As Boolean
            Get
                Return (Me.mPointer < If(Me.mTokens?.Count, 0))
            End Get
        End Property

        ''' <summary>カレント文字を返す。</summary>
        Public ReadOnly Property Current As TokenPosition
            Get
                Return If(Me.mPointer < If(Me.mTokens?.Count, 0), Me.mTokens(Me.mPointer), Nothing)
            End Get
        End Property

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="inputtkn">入力トークン。</param>
        Public Sub New(inputtkn As IEnumerable(Of TokenPosition))
            Me.mPointer = 0
            Me.mTokens = inputtkn.ToList()
        End Sub

        ''' <summary>カレント位置を移動させる。</summary>
        ''' <param name="moveAmount">移動量。</param>
        Public Sub Move(moveAmount As Integer)
            Me.mPointer += moveAmount
        End Sub

        ''' <summary>指定位置のトークンを取得する。</summary>
        ''' <param name="skipCount">スキップ数。</param>
        ''' <returns>トークン。</returns>
        Public Function Peek(skipCount As Integer) As TokenPosition
            Return If(Me.mPointer + skipCount < If(Me.mTokens?.Count, 0), Me.mTokens(Me.mPointer + skipCount), Nothing)
        End Function

    End Class

End Module
