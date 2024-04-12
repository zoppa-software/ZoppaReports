Option Strict On
Option Explicit On

Imports System.Text
Imports ZoppaReports.Exceptions
Imports ZoppaReports.Express
Imports ZoppaReports.Tokens
Imports ZoppaReports.ExTokens

Namespace Parser

    Partial Module ParserAnalysis

        ''' <summary>テンプレートの置き換えを実施します。</summary>
        ''' <param name="templateStr">テンプレート文字列。</param>
        ''' <param name="tokens">トークンリスト。</param>
        ''' <param name="parameter">環境値情報。</param>
        ''' <returns>置き換え結果。</returns>
        Friend Function Replase(templateStr As String,
                                tokens As List(Of TokenPosition),
                                parameter As Environments) As String
            ' 置き換えを実施
            Dim buffer As New StringBuilder()
            ReplaseQuery(templateStr, New TokenStream(tokens), parameter, buffer)

            ' 空白行を削除
            Return buffer.ToString()
        End Function

        ''' <summary>テンプレートの置き換えを実施します。</summary>
        ''' <param name="templateStr">テンプレート。</param>
        ''' <param name="reader">トークンリーダー。</param>
        ''' <param name="parameter">環境値情報。</param>
        Private Sub ReplaseQuery(templateStr As String,
                                 reader As TokenStream,
                                 parameter As Environments,
                                 buffer As StringBuilder)
            Do While reader.HasNext
                Dim tkn = reader.Current

                Select Case tkn.TokenType
                    Case GetType(QueryToken)
                        buffer.Append(tkn.ToString())
                        reader.Move(1)

                    Case GetType(ReplaseToken)
                        Dim rtoken = tkn.GetToken(Of ReplaseToken)()
                        If rtoken IsNot Nothing Then
                            Dim rval = parameter.GetValue(If(rtoken.Contents?.ToString(), ""))
                            Dim ans = GetRefValue(rval)
                            buffer.Append(ans)
                        End If
                        reader.Move(1)

                    Case GetType(IfToken)
                        reader.Move(1)
                        Dim ifTokens = CollectBlockToken(templateStr, reader, New Type() {GetType(IfToken)}, GetType(EndIfToken))
                        EvaluationIf(templateStr, tkn.Position, tkn.GetToken(Of IfToken)(), ifTokens, buffer, parameter)

                    Case GetType(ElseIfToken), GetType(ElseToken), GetType(EndIfToken)
                        Throw New ReportsAnalysisException($"ifが開始されていません。{vbCrLf}{templateStr}:{tkn.Position}")

                    Case GetType(ForToken)
                        reader.Move(1)
                        Dim forTokens = CollectBlockToken(templateStr, reader, New Type() {GetType(ForToken), GetType(ForEachToken)}, GetType(EndForToken))
                        EvaluationFor(templateStr, tkn.Position, tkn.GetToken(Of ForToken)(), forTokens, buffer, parameter)

                    Case GetType(ForEachToken)
                        reader.Move(1)
                        Dim forTokens = CollectBlockToken(templateStr, reader, New Type() {GetType(ForToken), GetType(ForEachToken)}, GetType(EndForToken))
                        EvaluationForEach(templateStr, tkn.Position, tkn.GetToken(Of ForEachToken)(), forTokens, buffer, parameter)

                    Case GetType(EndForToken)
                        Throw New ReportsAnalysisException($"forが開始されていません。{vbCrLf}{templateStr}:{tkn.Position}")

                    Case GetType(SelectToken)
                        reader.Move(1)
                        Dim selTokens = CollectBlockToken(templateStr, reader, New Type() {GetType(SelectToken)}, GetType(EndSelectToken))
                        EvaluationSelect(templateStr, tkn.Position, tkn.GetToken(Of SelectToken)(), selTokens, buffer, parameter)

                    Case GetType(CaseToken), GetType(EndSelectToken)
                        Throw New ReportsAnalysisException($"selectが開始されていません。{vbCrLf}{templateStr}:{tkn.Position}")
                End Select
            Loop
        End Sub

        ''' <summary>パラメータ値を参照して取得します。</summary>
        ''' <param name="refObj">参照オブジェクト。</param>
        ''' <returns>取得した値を表現する文字列。</returns>
        Private Function GetRefValue(refObj As Object) As String
            If TypeOf refObj Is String Then
                ' 文字列を取得
                Return refObj.ToString()

            ElseIf TypeOf refObj Is IEnumerable Then
                ' 列挙して値を取得
                Dim buf As New StringBuilder()
                For Each itm In CType(refObj, IEnumerable)
                    If buf.Length > 0 Then
                        buf.Append(", ")
                    End If
                    buf.Append(GetRefValue(itm))
                Next
                Return buf.ToString()

            ElseIf refObj Is Nothing Then
                ' null値を取得
                Return "null"
            Else
                Return refObj.ToString()
            End If
        End Function

        ''' <summary>ブロックトークンを集めます。</summary>
        ''' <param name="sqlQuery">テンプレート。</param>
        ''' <param name="reader">トークンリーダー。</param>
        ''' <param name="startTokenName">開始トークン。</param>
        ''' <param name="endTokenName">終了トークン。</param>
        ''' <returns>ブロックトークン。</returns>
        Private Function CollectBlockToken(templateStr As String,
                                           reader As TokenStream,
                                           startTokenName As Type(),
                                           endTokenName As Type) As List(Of TokenPosition)
            Dim res As New List(Of TokenPosition)()
            Dim startToken = reader.Current

            Dim nest As Integer = 0
            Do While reader.HasNext
                Dim tkn = reader.Current
                If startTokenName.Any(Function(t) tkn.TokenType Is t) Then
                    nest += 1
                ElseIf tkn.TokenType Is endTokenName Then
                    If nest = 0 Then
                        reader.Move(1)
                        Return res
                    Else
                        nest -= 1
                    End If
                End If
                res.Add(tkn)
                reader.Move(1)
            Loop

            Throw New ReportsAnalysisException($"フロックが閉じられていません。{vbCrLf}{templateStr}:{startToken.Position}")
        End Function

        ''' <summary>Ifを評価します。</summary>
        ''' <param name="templateStr">テンプレート。</param>
        ''' <param name="tempPos">評価位置。</param>
        ''' <param name="sifToken">Ifのトークンリスト。</param>
        ''' <param name="tokens">ブロック内のトークンリスト。</param>
        ''' <param name="buffer">結果バッファ。</param>
        ''' <param name="parameter">パラメータ。</param>
        Private Sub EvaluationIf(templateStr As String,
                                 tempPos As Integer,
                                 sifToken As IToken,
                                 tokens As List(Of TokenPosition),
                                 buffer As StringBuilder,
                                 parameter As Environments)
            Dim blocks As New List(Of (condition As IToken, block As List(Of TokenPosition))) From {
                (sifToken, New List(Of TokenPosition)())
            }

            '----------------
            ' トークン解析
            '----------------
            ' If、ElseIf、Elseブロックを集める
            Dim nest As Integer = 0
            For Each tkn In tokens
                Select Case tkn.TokenType
                    Case GetType(IfToken)
                        nest += 1
                        blocks(blocks.Count - 1).block.Add(tkn)

                    Case GetType(ElseIfToken)
                        If nest = 0 Then
                            blocks.Add((tkn.GetToken(Of ElseIfToken), New List(Of TokenPosition)()))
                        Else
                            blocks(blocks.Count - 1).block.Add(tkn)
                        End If

                    Case GetType(ElseToken)
                        If nest = 0 Then
                            blocks.Add((tkn.GetToken(Of ElseToken), New List(Of TokenPosition)()))
                        Else
                            blocks(blocks.Count - 1).block.Add(tkn)
                        End If

                    Case GetType(EndIfToken)
                        If nest = 0 Then
                            Exit For
                        Else
                            nest -= 1
                            blocks(blocks.Count - 1).block.Add(tkn)
                        End If

                    Case Else
                        blocks(blocks.Count - 1).block.Add(tkn)
                End Select
            Next

            '----------------
            ' 条件分岐実行
            '----------------
            Try
                ' If、ElseIf、Elseブロックを評価
                Dim lclbuf As New StringBuilder()
                For Each tkn In blocks
                    Select Case tkn.condition.TokenType
                        Case GetType(IfToken), GetType(ElseIfToken)
                            ' 条件を評価して真ならば、ブロックを出力
                            Dim ts As New TokenStream(CType(tkn.condition.Contents, List(Of TokenPosition)))
                            Dim answer = mMultiParser.Parser(ts).Executes(parameter).Contents
                            If TypeOf answer Is Boolean AndAlso CBool(answer) Then
                                ReplaseQuery(templateStr, New TokenStream(tkn.block), parameter, lclbuf)
                                Exit For
                            End If

                        Case GetType(ElseToken)
                            Dim tkns As New List(Of TokenPosition)(tkn.block)
                            ReplaseQuery(templateStr, New TokenStream(tkns), parameter, lclbuf)
                    End Select
                Next
                buffer.Append(lclbuf.ToString())
            Catch ex As Exception
                Throw New ReportsException($"ifの構文を間違えています。{vbCrLf}{templateStr}:{tempPos}", ex)
            End Try
        End Sub

        ''' <summary>Forを評価します。</summary>
        ''' <param name="templateStr">テンプレート。</param>
        ''' <param name="tempPos">評価位置。</param>
        ''' <param name="sforToken">Forのトークンリスト。</param>
        ''' <param name="tokens">ブロック内のトークンリスト。</param>
        ''' <param name="buffer">結果バッファ。</param>
        ''' <param name="parameter">パラメータ。</param>
        Private Sub EvaluationFor(templateStr As String,
                                  tempPos As Integer,
                                  sforToken As ForToken,
                                  tokens As List(Of TokenPosition),
                                  buffer As StringBuilder,
                                  parameter As Environments)
            '----------------
            ' トークン解析
            '----------------
            Dim ts As New TokenStream(CType(sforToken.Contents, List(Of TokenPosition)))

            ' カウンタ変数を取得
            Dim varName As String
            Dim varContents = ts.Current.GetToken(Of IdentToken)().Contents
            If varContents IsNot Nothing AndAlso Not parameter.IsDefainedName(varContents.ToString()) Then
                varName = varContents.ToString()
            Else
                Throw New ReportsAnalysisException($"変数が定義済みです。{vbCrLf}{templateStr}:{tempPos}")
            End If
            ts.Move(1)

            ' = トークンを取得
            If ts.Current.TokenType IsNot GetType(EqualToken) Then
                Throw New ReportsAnalysisException($"forの構文を間違えています。{vbCrLf}{templateStr}:{tempPos}")
            End If
            ts.Move(1)

            ' 開始値を取得
            Dim startToken = mMultiParser.Parser(ts)

            ' toトークンを取得
            Dim toToken = ts.Current.GetToken(Of IdentToken)()
            If toToken Is Nothing OrElse toToken.ToString().ToLower() <> "to" Then
                Throw New ReportsAnalysisException($"forの構文を間違えています。{vbCrLf}{templateStr}:{tempPos}")
            End If
            ts.Move(1)

            ' 終了値を取得
            Dim endToken = mMultiParser.Parser(ts)

            If ts.HasNext Then
                Throw New ReportsAnalysisException($"forの構文を間違えています。{vbCrLf}{templateStr}:{tempPos}")
            End If

            '----------------
            ' ループ実行
            '----------------
            Try
                For i As Integer = Convert.ToInt32(startToken.Executes(parameter).Contents) To Convert.ToInt32(endToken.Executes(parameter).Contents)
                    parameter.AddVariant(varName, i)

                    Dim lclbuf As New StringBuilder()
                    ReplaseQuery(templateStr, New TokenStream(tokens), parameter, lclbuf)
                    buffer.Append(lclbuf)

                    parameter.RemoveVariant(varName)
                Next
            Catch ex As Exception
                Throw New ReportsException($"forの構文を間違えています。{vbCrLf}{templateStr}:{tempPos}", ex)
            End Try
        End Sub

        ''' <summary>ForEachを評価します。</summary>
        ''' <param name="templateStr">テンプレート。</param>
        ''' <param name="tempPos">評価位置。</param>
        ''' <param name="sforToken">Forのトークンリスト。</param>
        ''' <param name="tokens">ブロック内のトークンリスト。</param>
        ''' <param name="buffer">結果バッファ。</param>
        ''' <param name="parameter">パラメータ。</param>
        Private Sub EvaluationForEach(templateStr As String,
                                      tempPos As Integer,
                                      sforToken As ForEachToken,
                                      tokens As List(Of TokenPosition),
                                      buffer As StringBuilder,
                                      parameter As Environments)
            '----------------
            ' トークン解析
            '----------------
            Dim ts As New TokenStream(CType(sforToken.Contents, List(Of TokenPosition)))

            ' カウンタ変数を取得
            Dim varName As String
            Dim varContents = ts.Current.GetToken(Of IdentToken)().Contents
            If varContents IsNot Nothing AndAlso Not parameter.IsDefainedName(varContents.ToString()) Then
                varName = varContents.ToString()
            Else
                Throw New ReportsAnalysisException($"変数が定義済みです。{vbCrLf}{templateStr}:{tempPos}")
            End If
            ts.Move(1)

            ' in、:トークンを取得
            Dim toToken = ts.Current.GetToken(Of IdentToken)()
            If toToken Is Nothing OrElse toToken.ToString().ToLower() <> "in" Then
                Throw New ReportsAnalysisException($"foreachの構文を間違えています。{vbCrLf}{templateStr}:{tempPos}")
            End If
            ts.Move(1)

            ' コレクションを取得
            Dim collectionToken = mMultiParser.Parser(ts)
            Dim collection = TryCast(collectionToken.Executes(parameter).Contents, IEnumerable)
            If collection Is Nothing Then
                Throw New ReportsAnalysisException($"foreachの構文を間違えています。{vbCrLf}{templateStr}:{tempPos}")
            End If

            If ts.HasNext Then
                Throw New ReportsAnalysisException($"foreachの構文を間違えています。{vbCrLf}{templateStr}:{tempPos}")
            End If

            '----------------
            ' ループ実行
            '----------------
            Try
                For Each v In collection
                    parameter.AddVariant(varName, v)

                    Dim lclbuf As New StringBuilder()
                    ReplaseQuery(templateStr, New TokenStream(tokens), parameter, lclbuf)
                    buffer.Append(lclbuf)

                    parameter.RemoveVariant(varName)
                Next
            Catch ex As Exception
                Throw New ReportsException($"foreachの構文を間違えています。{vbCrLf}{templateStr}:{tempPos}", ex)
            End Try
        End Sub

        ''' <summary>Selectを評価します。</summary>
        ''' <param name="templateStr">テンプレート。</param>
        ''' <param name="tempPos">評価位置。</param>
        ''' <param name="sselToken">Ifのトークンリスト。</param>
        ''' <param name="tokens">ブロック内のトークンリスト。</param>
        ''' <param name="buffer">結果バッファ。</param>
        ''' <param name="parameter">パラメータ。</param>
        Private Sub EvaluationSelect(templateStr As String,
                                     tempPos As Integer,
                                     sselToken As IToken,
                                     tokens As List(Of TokenPosition),
                                     buffer As StringBuilder,
                                     parameter As Environments)
            Dim blocks As New List(Of (condition As IToken, block As List(Of TokenPosition))) From {
                (sselToken, New List(Of TokenPosition)())
            }

            '----------------
            ' トークン解析
            '----------------
            ' Select、Case、Elseブロックを集める
            Dim nest As Integer = 0
            Dim isSkip As Boolean = True
            For Each tkn In tokens
                If isSkip AndAlso tkn.TokenType Is GetType(QueryToken) AndAlso tkn.GetToken(Of QueryToken).Contents?.ToString().Trim() = "" Then
                    Continue For
                End If
                isSkip = False

                Select Case tkn.TokenType
                    Case GetType(SelectToken)
                        nest += 1
                        blocks(blocks.Count - 1).block.Add(tkn)

                    Case GetType(CaseToken)
                        If nest = 0 Then
                            blocks.Add((tkn.GetToken(Of CaseToken), New List(Of TokenPosition)()))
                        Else
                            blocks(blocks.Count - 1).block.Add(tkn)
                        End If

                    Case GetType(ElseToken)
                        If nest = 0 Then
                            blocks.Add((tkn.GetToken(Of ElseToken), New List(Of TokenPosition)()))
                        Else
                            blocks(blocks.Count - 1).block.Add(tkn)
                        End If

                    Case GetType(EndSelectToken)
                        If nest = 0 Then
                            Exit For
                        Else
                            nest -= 1
                            blocks(blocks.Count - 1).block.Add(tkn)
                        End If

                    Case Else
                        blocks(blocks.Count - 1).block.Add(tkn)
                End Select
            Next

            '----------------
            ' Select実行
            '----------------
            Try
                ' Select、Case、Elseブロックを評価
                Dim lclbuf As New StringBuilder()

                Dim selectTs As New TokenStream(CType(blocks(0).condition.Contents, List(Of TokenPosition)))
                Dim selectVal = mMultiParser.Parser(selectTs).Executes(parameter).Contents

                For i As Integer = 1 To blocks.Count - 1
                    Dim tkn = blocks(i)

                    Select Case tkn.condition.TokenType
                        Case GetType(CaseToken)
                            ' 条件を評価して真ならば、ブロックを出力
                            Dim caseTs As New TokenStream(CType(tkn.condition.Contents, List(Of TokenPosition)))
                            Dim caseVal = mMultiParser.Parser(caseTs).Executes(parameter).Contents
                            If caseVal?.Equals(selectVal) Then
                                Dim tkns As New List(Of TokenPosition)(tkn.block)
                                ReplaseQuery(templateStr, New TokenStream(tkns), parameter, lclbuf)
                                Exit For
                            End If

                        Case GetType(ElseToken)
                            Dim tkns As New List(Of TokenPosition)(tkn.block)
                            ReplaseQuery(templateStr, New TokenStream(tkns), parameter, lclbuf)
                    End Select
                Next

                buffer.Append(lclbuf.ToString())

            Catch ex As Exception
                Throw New ReportsException($"selectの構文を間違えています。{vbCrLf}{templateStr}:{tempPos}", ex)
            End Try
        End Sub

    End Module

End Namespace
