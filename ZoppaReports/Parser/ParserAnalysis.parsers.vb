Option Strict On
Option Explicit On

Imports ZoppaReports.Exceptions
Imports ZoppaReports.Express
Imports ZoppaReports.Tokens

Partial Module ParserAnalysis

    ''' <summary>解析インターフェイス。</summary>
    Private Interface IParser

        ''' <summary>解析を実行する。</summary>
        ''' <param name="reader">入力トークンストリーム。</param>
        ''' <returns>解析結果。</returns>
        Function Parser(reader As TokenStream) As IExpression

    End Interface

    ''' <summary>括弧解析。</summary>
    Private NotInheritable Class ParenParser
        Implements IParser

        ''' <summary>次のパーサーを設定、取得する。</summary>
        Friend Property NextParser() As IParser

        ''' <summary>解析を実行する。</summary>
        ''' <param name="reader">入力トークンストリーム。</param>
        ''' <returns>解析結果。</returns>
        Public Function Parser(reader As TokenStream) As IExpression Implements IParser.Parser
            Dim tkn = reader.Current
            If tkn.TokenType Is GetType(LParenToken) Then
                reader.Move(1)
                Return CreateParenExpress(Of LParenToken, RParenToken)(reader, Me.NextParser)
            Else
                Return Me.NextParser.Parser(reader)
            End If
        End Function

    End Class

    ''' <summary>論理解析。</summary>
    Private NotInheritable Class LogicalParser
        Implements IParser

        ''' <summary>次のパーサーを設定、取得する。</summary>
        Friend Property NextParser() As IParser

        ''' <summary>解析を実行する。</summary>
        ''' <param name="reader">入力トークンストリーム。</param>
        ''' <returns>解析結果。</returns>
        Public Function Parser(reader As TokenStream) As IExpression Implements IParser.Parser
            Dim tml = Me.NextParser.Parser(reader)

            Do While reader.HasNext
                Dim ope = reader.Current
                Select Case ope.TokenType
                    Case GetType(AndToken)
                        reader.Move(1)
                        tml = New AndExpress(tml, Me.NextParser.Parser(reader))

                    Case GetType(OrToken)
                        reader.Move(1)
                        tml = New OrExpress(tml, Me.NextParser.Parser(reader))

                    Case Else
                        Exit Do
                End Select
            Loop

            Return tml
        End Function

    End Class

    ''' <summary>比較解析。</summary>
    Private NotInheritable Class ComparisonParser
        Implements IParser

        ''' <summary>次のパーサーを設定、取得する。</summary>
        Friend Property NextParser() As IParser

        ''' <summary>解析を実行する。</summary>
        ''' <param name="reader">入力トークンストリーム。</param>
        ''' <returns>解析結果。</returns>
        Public Function Parser(reader As TokenStream) As IExpression Implements IParser.Parser
            Dim tml = Me.NextParser.Parser(reader)

            If reader.HasNext Then
                Dim ope = reader.Current
                Select Case ope.TokenType
                    Case GetType(EqualToken)
                        reader.Move(1)
                        tml = New EqualExpress(tml, Me.NextParser.Parser(reader))

                    Case GetType(NotEqualToken)
                        reader.Move(1)
                        tml = New NotEqualExpress(tml, Me.NextParser.Parser(reader))

                    Case GetType(GreaterToken)
                        reader.Move(1)
                        tml = New GreaterExpress(tml, Me.NextParser.Parser(reader))

                    Case GetType(GreaterEqualToken)
                        reader.Move(1)
                        tml = New GreaterEqualExpress(tml, Me.NextParser.Parser(reader))

                    Case GetType(LessToken)
                        reader.Move(1)
                        tml = New LessExpress(tml, Me.NextParser.Parser(reader))

                    Case GetType(LessEqualToken)
                        reader.Move(1)
                        tml = New LessEqualExpress(tml, Me.NextParser.Parser(reader))
                End Select
            End If

            Return tml
        End Function

    End Class

    ''' <summary>加算、減算解析。</summary>
    Private NotInheritable Class AddOrSubParser
        Implements IParser

        ''' <summary>次のパーサーを設定、取得する。</summary>
        Friend Property NextParser() As IParser

        ''' <summary>解析を実行する。</summary>
        ''' <param name="reader">入力トークンストリーム。</param>
        ''' <returns>解析結果。</returns>
        Public Function Parser(reader As TokenStream) As IExpression Implements IParser.Parser
            Dim tml = Me.NextParser.Parser(reader)

            Do While reader.HasNext
                Dim ope = reader.Current
                Select Case ope.TokenType
                    Case GetType(PlusToken)
                        reader.Move(1)
                        tml = New PlusExpress(tml, Me.NextParser.Parser(reader))

                    Case GetType(MinusToken)
                        reader.Move(1)
                        tml = New MinusExpress(tml, Me.NextParser.Parser(reader))

                    Case Else
                        Exit Do
                End Select
            Loop

            Return tml
        End Function

    End Class

    ''' <summary>乗算、除算解析。</summary>
    Private NotInheritable Class MultiOrDivParser
        Implements IParser

        ''' <summary>次のパーサーを設定、取得する。</summary>
        Friend Property NextParser() As IParser

        ''' <summary>解析を実行する。</summary>
        ''' <param name="reader">入力トークンストリーム。</param>
        ''' <returns>解析結果。</returns>
        Public Function Parser(reader As TokenStream) As IExpression Implements IParser.Parser
            Dim tml = Me.NextParser.Parser(reader)

            Do While reader.HasNext
                Dim ope = reader.Current
                Select Case ope.TokenType
                    Case GetType(MultiToken)
                        reader.Move(1)
                        tml = New MultiExpress(tml, Me.NextParser.Parser(reader))

                    Case GetType(DivToken)
                        reader.Move(1)
                        tml = New DivExpress(tml, Me.NextParser.Parser(reader))

                    Case Else
                        Exit Do
                End Select
            Loop

            Return tml
        End Function
    End Class

    ''' <summary>参照解析。</summary>
    Private NotInheritable Class ReferenceParser
        Implements IParser

        ''' <summary>次のパーサーを設定、取得する。</summary>
        Friend Property NextParser() As IParser

        ''' <summary>解析を実行する。</summary>
        ''' <param name="reader">入力トークンストリーム。</param>
        ''' <returns>解析結果。</returns> 
        Public Function Parser(reader As TokenStream) As IExpression Implements IParser.Parser
            Dim tml = Me.NextParser.Parser(reader)

            Do While reader.HasNext
                Dim ope = reader.Current
                Select Case ope.TokenType
                    Case GetType(PeriodToken)
                        reader.Move(1)
                        tml = New ReferenceExpress(tml, Me.NextParser.Parser(reader))

                    Case Else
                        Exit Do
                End Select
            Loop

            Return tml
        End Function

    End Class

    ''' <summary>要素解析。</summary>
    Private NotInheritable Class FactorParser
        Implements IParser

        ''' <summary>次のパーサーを設定、取得する。</summary>
        Friend Property NextParser() As IParser

        ''' <summary>解析を実行する。</summary>
        ''' <param name="reader">入力トークンストリーム。</param>
        ''' <returns>解析結果。</returns>
        Public Function Parser(reader As TokenStream) As IExpression Implements IParser.Parser
            Dim tkn = reader.Current

            Select Case tkn.TokenType
                Case GetType(LParenToken)
                    reader.Move(1)
                    Return CreateParenExpress(Of LParenToken, RParenToken)(reader, Me.NextParser)

                Case GetType(IdentToken)
                    reader.Move(1)
                    If reader.Current.TokenType Is GetType(LBracketToken) Then
                        reader.Move(1)
                        Dim numTkn = CreateParenExpress(Of LBracketToken, RBracketToken)(reader, Me.NextParser)
                        Return New ArrayValueExpress(tkn.GetToken(Of IdentToken)(), numTkn)
                    Else
                        Return New ValueExpress(tkn.GetToken(Of IdentToken)())
                    End If

                Case GetType(NumberToken), GetType(StringToken),
                     GetType(TrueToken), GetType(FalseToken),
                     GetType(NullToken), GetType(ObjectToken)
                    '     GetType(QueryToken), GetType(ReplaseToken),
                    reader.Move(1)
                    Return New ValueExpress(tkn.GetToken(Of IToken)())

                Case GetType(PlusToken), GetType(MinusToken), GetType(NotToken)
                    reader.Move(1)
                    Dim nxtExper = Me.Parser(reader)
                    If TypeOf nxtExper Is ValueExpress Then
                        Return New UnaryExpress(tkn.GetToken(Of IToken)(), nxtExper)
                    Else
                        Throw New ReportsAnalysisException($"前置き演算子{tkn}が値の前に配置していません")
                    End If

                Case Else
                    Throw New ReportsAnalysisException("Factor要素の解析に失敗")
            End Select
        End Function

    End Class

End Module
