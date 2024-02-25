Option Strict On
Option Explicit On

Imports System.Runtime.CompilerServices
Imports System.Text
Imports ZoppaReports.Exceptions
Imports ZoppaReports.ExTokens
Imports ZoppaReports.Tokens

Namespace Lexical

    Partial Module LexicalAnalysis

        ''' <summary>トークン種別。</summary>
        Private Enum TKN_TYPE

            ''' <summary>クエリトークン。</summary>
            QUERY_TKN

            ''' <summary>コードトークン。</summary>
            CODE_TKN

            ''' <summary>置き換えトークン。</summary>
            REPLASE_TKN

        End Enum

        ''' <summary>文字列をトークン解析します。</summary>
        ''' <param name="command">文字列。</param>
        ''' <returns>トークンリスト。</returns>
        Public Function SplitRepToken(command As String) As List(Of TokenPosition)
            Dim tokens As New List(Of TokenPosition)()

            Dim reader = New StringPtr(command)

            Dim buffer As New StringBuilder()
            Dim tknType = TKN_TYPE.QUERY_TKN
            Dim pos As Integer = 0
            Do While reader.HasNext
                Dim c = reader.Current()

                Select Case c
                    Case "{"c
                        tokens.AddRange(buffer.CreateQueryTokens(pos))
                        pos = reader.CurrentPosition
                        reader.Move(1)
                        tknType = TKN_TYPE.CODE_TKN

                    Case "}"c
                        Select Case tknType
                            Case TKN_TYPE.QUERY_TKN
                                Throw New ReportsAnalysisException("}に対応する{を入力してください")
                            Case TKN_TYPE.CODE_TKN
                                tokens.AddIfNull(buffer.CreateCodeToken(), pos)
                            Case TKN_TYPE.REPLASE_TKN
                                tokens.AddIfNull(buffer.CreateReplaseToken(), pos)
                        End Select
                        pos = reader.CurrentPosition
                        reader.Move(1)
                        tknType = TKN_TYPE.QUERY_TKN

                    Case "#"c
                        If reader.NestChar(1) = "{"c Then
                            tokens.AddRange(buffer.CreateQueryTokens(pos))
                            pos = reader.CurrentPosition
                            reader.Move(2)
                            tknType = TKN_TYPE.REPLASE_TKN
                        Else
                            buffer.Append(c)
                            reader.Move(1)
                        End If

                    Case "\"c
                        If reader.NestChar(1) = "{"c OrElse reader.NestChar(1) = "}"c Then
                            reader.Move(1)
                            buffer.Append(reader.Current)
                            reader.Move(1)
                        Else
                            buffer.Append(c)
                            reader.Move(1)
                        End If

                    Case Else
                        buffer.Append(c)
                        reader.Move(1)
                End Select
            Loop

            Select Case tknType
                Case TKN_TYPE.QUERY_TKN
                    tokens.AddRange(buffer.CreateQueryTokens(pos))
                Case TKN_TYPE.CODE_TKN
                    Throw New ReportsAnalysisException("コードトークンが閉じられていません")
                Case TKN_TYPE.REPLASE_TKN
                    Throw New ReportsAnalysisException("置き換えトークンが閉じられていません")
            End Select

            Return tokens
        End Function

        ''' <summary>トークンリストにトークンを追加します(Nullを除く)</summary>
        ''' <param name="tokens">トークンリスト。</param>
        ''' <param name="token">追加するトークン。</param>
        ''' <param name="pos">トークン位置。</param>
        <Extension>
        Private Sub AddIfNull(tokens As List(Of TokenPosition), token As IToken, pos As Integer)
            If token IsNot Nothing Then
                tokens.Add(New TokenPosition(token, pos))
            End If
        End Sub

        ''' <summary>直接出力するクエリトークンを生成します。</summary>
        ''' <param name="buffer">文字列バッファ。</param>
        ''' <param name="startPos">トークンの最初の文字位置。</param>
        ''' <returns>直接出力するクエリトークンを生成します。</returns>
        <Extension>
        Private Function CreateQueryTokens(buffer As StringBuilder, startPos As Integer) As List(Of TokenPosition)
            Dim res As New List(Of TokenPosition)
            If buffer.Length > 0 Then
                res.Add(New TokenPosition(New QueryToken(buffer.ToString()), startPos))
            End If
            buffer.Clear()
            Return res
        End Function

        ''' <summary>置き換えトークンを生成します。</summary>
        ''' <param name="buffer">文字列バッファ。</param>
        ''' <returns>置き換えトークン。</returns>
        <Extension>
        Private Function CreateReplaseToken(buffer As StringBuilder) As IToken
            Dim repStr = buffer.ToString().Trim()
            buffer.Clear()
            Return If(repStr.Length > 0, New ReplaseToken(repStr), Nothing)
        End Function

        ''' <summary>コードトークンリストを生成します。</summary>
        ''' <param name="buffer">文字列バッファ。</param>
        ''' <returns>コードトークン</returns>
        <Extension>
        Private Function CreateCodeToken(buffer As StringBuilder) As IToken
            Dim codeStr = buffer.ToString().Trim()
            buffer.Clear()
            Dim lowStr = If(codeStr.Length > 10, codeStr.Substring(0, 10), codeStr).ToLower()

            If lowStr.StartsWith("for ") Then
                Return New ForToken(SplitToken(codeStr.Substring(4)))
            ElseIf lowStr.StartsWith("end for") Then
                Return EndForToken.Value
            ElseIf lowStr.StartsWith("/for") Then
                Return EndForToken.Value
            End If

            'If lowStr.StartsWith("if ") Then
            '    Return New IfToken(SplitToken(codeStr.Substring(3)))
            'ElseIf lowStr.StartsWith("else if ") Then
            '    Return New ElseIfToken(SplitToken(codeStr.Substring(8)))
            'ElseIf lowStr.StartsWith("else") Then
            '    Return ElseToken.Value
            'ElseIf lowStr.StartsWith("end if") Then
            '    Return EndIfToken.Value
            'ElseIf lowStr.StartsWith("/if") Then
            '    Return EndIfToken.Value
            'ElseIf lowStr.StartsWith("for each ") Then
            '    Return New ForEachToken(SplitToken(codeStr.Substring(9)))
            'ElseIf lowStr.StartsWith("foreach ") Then
            '    Return New ForEachToken(SplitToken(codeStr.Substring(8)))
            'ElseIf lowStr.StartsWith("end for") Then
            '    Return EndForToken.Value
            'ElseIf lowStr.StartsWith("/for") Then
            '    Return EndForToken.Value
            'ElseIf lowStr.StartsWith("select ") Then
            '    Return New SelectToken(SplitToken(codeStr.Substring(7)))
            'ElseIf lowStr.StartsWith("case ") Then
            '    Return New CaseToken(SplitToken(codeStr.Substring(5)))
            'ElseIf lowStr.StartsWith("end select") Then
            '    Return EndSelectToken.Value
            'ElseIf lowStr.StartsWith("/select") Then
            '    Return EndSelectToken.Value
            'End If
            Return Nothing
        End Function

    End Module

End Namespace
