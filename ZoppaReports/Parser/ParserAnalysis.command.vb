﻿Option Strict On
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
            ' SQLの置き換えを実施
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

                    Case GetType(ForToken)
                        reader.Move(1)
                        Dim forTokens = CollectBlockToken(templateStr, reader, New Type() {GetType(ForToken)}, GetType(EndForToken))
                        EvaluationFor(templateStr, tkn.Position, tkn.GetToken(Of ForToken)(), forTokens, buffer, parameter)

                    Case GetType(EndForToken)
                        Throw New ReportsAnalysisException($"forが開始されていません。{vbCrLf}{templateStr}:{tkn.Position}")

                        'Case GetType(IfToken)
                        '    reader.Move(1)
                        '    Dim ifTokens = CollectBlockToken(sqlQuery, reader, GetType(IfToken), GetType(EndIfToken))
                        '    EvaluationIf(sqlQuery, tkn.GetToken(Of IfToken)(), ifTokens, buffer, parameter)

                        'Case GetType(ElseIfToken), GetType(ElseToken), GetType(EndIfToken)
                        '    Throw New DSqlAnalysisException($"ifが開始されていません。{vbCrLf}{sqlQuery}:{tkn.Position}")

                        'Case GetType(ForEachToken)
                        '    reader.Move(1)
                        '    Dim forTokens = CollectBlockToken(sqlQuery, reader, GetType(ForEachToken), GetType(EndForToken))
                        '    EvaluationFor(sqlQuery, tkn.GetToken(Of ForEachToken)(), forTokens, buffer, parameter)

                        'Case GetType(EndForToken)
                        '    Throw New DSqlAnalysisException($"forが開始されていません。{vbCrLf}{sqlQuery}:{tkn.Position}")

                        'Case GetType(SelectToken)
                        '    reader.Move(1)
                        '    Dim selTokens = CollectBlockToken(sqlQuery, reader, GetType(SelectToken), GetType(EndSelectToken))
                        '    EvaluationSelect(sqlQuery, tkn.GetToken(Of SelectToken)(), selTokens, buffer, parameter)

                        'Case GetType(CaseToken), GetType(EndSelectToken)
                        '    Throw New DSqlAnalysisException($"selectが開始されていません。{vbCrLf}{sqlQuery}:{tkn.Position}")

                        'Case GetType(TrimToken)
                        '    reader.Move(1)
                        '    Dim trimTokens = CollectBlockToken(sqlQuery, reader, GetType(TrimToken), GetType(EndTrimToken))
                        '    EvaluationTrim(sqlQuery, tkn.GetToken(Of TrimToken)(), trimTokens, buffer, parameter)

                        'Case GetType(EndTrimToken)
                        '    Throw New DSqlAnalysisException($"trimが開始されていません。{vbCrLf}{sqlQuery}:{tkn.Position}")
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
                    buffer.Append(lclbuf.ToString())

                    parameter.RemoveVariant(varName)
                Next
            Catch ex As Exception
                Throw New ReportsException($"forの構文を間違えています。{vbCrLf}{templateStr}:{tempPos}", ex)
            End Try
        End Sub

    End Module

End Namespace