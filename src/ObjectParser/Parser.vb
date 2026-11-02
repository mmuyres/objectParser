Imports System.IO

Namespace ObjectParser

    ''' <summary>One query result: the matched text and where it sits in the parser's text.</summary>
    Public Structure Match
        Public Sub New(start As Integer, value As String)
            Me.Start = start
            Me.Value = value
        End Sub

        ''' <summary>Offset of <see cref="Value"/> in <see cref="Parser.Text"/>.</summary>
        Public ReadOnly Property Start As Integer
        Public ReadOnly Property Value As String
        Public ReadOnly Property Length As Integer
            Get
                Return Value.Length
            End Get
        End Property

        Public Overrides Function ToString() As String
            Return $"@{Start}: {Value}"
        End Function
    End Structure

    ''' <summary>
    ''' Evaluates paths against a text buffer using a <see cref="TokenMap"/>. There is no DOM: each
    ''' path segment's prefix selects a token pair, and the pair's delimiters are searched for in
    ''' the results of the previous segment.
    ''' </summary>
    Public NotInheritable Class Parser

        Private ReadOnly _map As TokenMap
        Private ReadOnly _syntax As ParserSyntax
        Private _position As Integer
        Private _previousMultiPath As String = ""

        Public Sub New(map As TokenMap)
            Me.New(map, "")
        End Sub

        Public Sub New(map As TokenMap, text As String)
            _map = If(map, TokenMap.BuiltIn("xml"))
            _syntax = _map.Syntax
            Me.Text = text
        End Sub

        Public ReadOnly Property Map As TokenMap
            Get
                Return _map
            End Get
        End Property

        ''' <summary>The text being queried. <see cref="Delete"/> and <see cref="Clean"/> change it.</summary>
        Public Property Text As String
            Get
                Return _text
            End Get
            Set(value As String)
                _text = If(value, "")
            End Set
        End Property
        Private _text As String = ""

        ''' <summary>Where trace output goes when <see cref="Debug"/> or <see cref="Verbose"/> is on.</summary>
        Public Property Log As TextWriter
        Public Property Debug As Boolean
        Public Property Verbose As Boolean

        Public ReadOnly Property Statistics As New ParseStatistics()

        ''' <summary>The matches from the most recent query.</summary>
        Public ReadOnly Property LastMatches As IReadOnlyList(Of Match) = Array.Empty(Of Match)()

#Region "Public API"

        ''' <summary>Evaluates <paramref name="path"/> and returns the matched text, in document order.</summary>
        Public Function Query(path As String) As IReadOnlyList(Of String)
            Return QueryMatches(path).Select(Function(m) m.Value).ToList()
        End Function

        ''' <summary>Evaluates <paramref name="path"/> and returns matches with their positions.</summary>
        Public Function QueryMatches(path As String) As IReadOnlyList(Of Match)
            Dim result = Evaluate(path)
            _LastMatches = result
            Return result
        End Function

        ''' <summary>Evaluates several paths and concatenates their results, in the order given.</summary>
        Public Function QueryAll(paths As IEnumerable(Of String)) As IReadOnlyList(Of String)
            Dim all = New List(Of Match)
            For Each p In paths
                all.AddRange(Evaluate(p))
            Next
            _LastMatches = all
            Return all.Select(Function(m) m.Value).ToList()
        End Function

        ''' <summary>
        ''' Finds the text between two ad-hoc tokens (token grammar applies; <c>~</c> means a double quote).
        ''' </summary>
        Public Function FindBetween(openToken As String, closeToken As String, Optional io As String = "I") As IReadOnlyList(Of String)
            Dim definition = New TokenDefinition("", openToken.Replace(_syntax.QuoteToken, """"),
                                                 closeToken.Replace(_syntax.QuoteToken, """"), io)
            Dim pair = New TokenPair(_syntax, definition)
            _position = 0
            Dim result = FindInText(Text, 0, "", pair, Nothing)
            _LastMatches = result
            Return result.Select(Function(m) m.Value).ToList()
        End Function

        ''' <summary>
        ''' Removes every match of <paramref name="path"/> from <see cref="Text"/> and returns the new text.
        ''' </summary>
        ''' <param name="replacement">Text to put in place of each match (default: nothing).</param>
        ''' <param name="keepPath">If given, each match is replaced by the results of this path evaluated inside it.</param>
        Public Function Delete(path As String, Optional replacement As String = Nothing, Optional keepPath As String = Nothing) As String
            Dim matches = QueryMatches(path)
            Text = Splice(Text, matches, Function(m)
                                             If Not String.IsNullOrEmpty(keepPath) Then
                                                 Dim inner = CreateChild(m.Value)
                                                 Return String.Join(Environment.NewLine, inner.Query(keepPath))
                                             End If
                                             Return If(replacement, "")
                                         End Function)
            Return Text
        End Function

        ''' <summary>Deletes the matches of each path in turn (used to strip noise before querying).</summary>
        Public Sub Clean(paths As IEnumerable(Of String))
            For Each p In paths
                Delete(p)
            Next
        End Sub

        Private Shared Function Splice(source As String, matches As IEnumerable(Of Match), replace As Func(Of Match, String)) As String
            Dim ordered = matches.Where(Function(m) m.Length > 0).
                OrderBy(Function(m) m.Start).ThenByDescending(Function(m) m.Length).ToList()
            Dim sb = New Text.StringBuilder()
            Dim cursor = 0
            For Each m In ordered
                If m.Start < cursor Then Continue For ' nested inside a match already removed
                sb.Append(source, cursor, m.Start - cursor)
                sb.Append(replace(m))
                cursor = m.Start + m.Length
            Next
            sb.Append(source, cursor, source.Length - cursor)
            Return sb.ToString()
        End Function

#End Region

#Region "Path evaluation"

        Private Function CreateChild(text As String) As Parser
            Dim child = New Parser(_map, text) With {
                .Log = Log, .Debug = Debug, .Verbose = Verbose,
                ._previousMultiPath = _previousMultiPath
            }
            Return child
        End Function

        Private Function Evaluate(path As String) As List(Of Match)
            Dim expr = PathExpression.Parse(path, _map.Prefixes, _syntax)
            If Verbose OrElse Debug Then Trace($"Path: {path}")

            If expr.SubPath.Length > 0 Then
                Dim child = CreateChild(Text)
                Dim values = child.Query(expr.SubPath)
                Statistics.Add(child.Statistics, includeSegments:=False)
                expr.ApplySubPathResult(String.Join(vbLf, values).Trim())
            End If

            Dim blocks As List(Of Match) = Nothing
            Dim previousPair As TokenPair = Nothing

            For Each segment In expr.Segments
                If Not _map.Contains(segment.Prefix) Then
                    Trace($"No token pair for prefix [{segment.Prefix}] in segment [{segment.Text}]; segment skipped.")
                    Continue For
                End If
                Dim started = Stopwatch.GetTimestamp()

                Dim names = expr.MultiPathNames(segment, _previousMultiPath)
                If names Is Nothing Then names = {segment.Field}
                If Debug Then Trace($"Segment: [{segment.Text}] prefix [{segment.Prefix}] field [{String.Join(",", names)}]")

                ' after the first segment, skip past the parent's own opening tag so it is not matched again
                Dim startOffset = 0
                If previousPair IsNot Nothing AndAlso previousPair.IO = "O" Then
                    startOffset = previousPair.OpenToken.FirstTagValue.Length
                End If

                Dim inputs = If(blocks, New List(Of Match) From {New Match(0, Text)})
                Dim output = New List(Of Match)
                For Each block In inputs
                    For Each name In names
                        Dim pair = _map.CreatePair(segment.Prefix)
                        _position = If(blocks Is Nothing, -1, startOffset)
                        output.AddRange(FindInText(block.Value, block.Start, name, pair, segment))
                        previousPair = pair
                    Next
                Next
                blocks = output
                Statistics.AddSegment(segment.Text, Stopwatch.GetElapsedTime(started).TotalMilliseconds)
            Next

            Return If(blocks, New List(Of Match))
        End Function

#End Region

#Region "Token search"

        ''' <summary>Finds every match of <paramref name="pair"/> in <paramref name="text"/>, then applies the segment filter.</summary>
        Private Function FindInText(text As String, offset As Integer, field As String, pair As TokenPair, segment As PathSegment) As List(Of Match)
            Dim started = Stopwatch.GetTimestamp()
            Dim results = New List(Of Match)
            Dim open = pair.OpenToken
            Dim close = pair.CloseToken
            pair.Reset()
            open.ClearPositions()
            close.ClearPositions()
            open.FieldValue = field
            close.FieldValue = field
            pair.IsWildCarding = False

            _position -= 1
            Dim lastOpenPos = -1
            Dim retrying = False
            Do
                If pair.IsWildCarding Then
                    open.FieldValue = ""
                    close.FieldValue = ""
                    If Not retrying Then _position -= 1
                End If
                retrying = False
                close.Reset()

                If Not ProcessToken(open, text) Then
                    If pair.IsCurrentResult Then results.Add(New Match(offset, text))
                    Exit Do
                End If
                If open.FirstTagPos <= lastOpenPos Then Exit Do ' no progress: same open match again
                lastOpenPos = open.FirstTagPos

                close.EvalCondition(open)
                If Not ProcessToken(close, text) Then
                    ' this candidate has no close; try the next open match
                    _position = open.FirstTagPos + 1
                    retrying = True
                    Continue Do
                End If

                If close.IsFieldResults Then
                    If open.FieldValue.Length > 0 Then results.Add(New Match(offset + open.FieldStart, open.FieldValue))
                    Continue Do
                End If

                Dim start As Integer, length As Integer
                If close.IsOuterResults Then
                    start = open.OuterPos
                    length = close.OuterPos - open.OuterPos
                Else
                    start = open.InnerPos
                    length = close.InnerPos - open.InnerPos
                End If
                If length < 0 Then Exit Do
                If start >= 0 AndAlso start + length <= text.Length Then
                    results.Add(New Match(offset + start, text.Substring(start, length)))
                    If Debug Then Trace($"  {(If(close.IsOuterResults, "outer", "inner"))} match @{offset + start} len {length}")
                End If
            Loop
            Statistics.ParseMilliseconds += Stopwatch.GetElapsedTime(started).TotalMilliseconds

            If segment IsNot Nothing AndAlso segment.HasFilter Then
                results = ApplyFilter(segment.Filter, text, results)
            End If
            Return results
        End Function

        ''' <summary>
        ''' Advances one token through <paramref name="text"/> from the current position. Returns True
        ''' when the token (and any jump or conditional part) matched.
        ''' </summary>
        Private Function ProcessToken(token As ParserToken, text As String) As Boolean
            If token.Parent.IsCurrentResult Then
                _position = -1
                Return False
            End If
            If token.Skip Then Return True ' the open token already consumed this close token

            If _position < 0 Then _position = 0
            If token.IsBalanced Then Return MatchBalanced(token, text)

            If _map.WildCardSupportEnabled AndAlso token.Role = TokenRole.Open AndAlso token.FieldIsWild Then
                If Not DeriveWildcardField(token, text) Then Return False
            End If

            token.FirstTagPos = If(text.Length > 0, FindFirstTag(token, text, _position), 0)
            If token.HadError OrElse token.FirstPosEnd > text.Length Then
                If Debug Then Trace($"  {token}: first tag [{token.FirstTagValue}] not found")
                Return False
            End If

            _position = token.FirstPosEnd
            If token.FirstPosEnd + token.NextPrevCharsSize > text.Length Then
                token.NextPrevCharsSize = text.Length - token.FirstPosEnd
            End If
            token.FirstTagNextChars = text.Substring(token.FirstPosEnd, token.NextPrevCharsSize)
            token.NextTagPrevChars = ""
            If Debug Then Trace($"  {token}: first tag [{token.FirstTagValue}] @{token.FirstTagPos} next chars [{token.FirstTagNextChars}]")

            If token.IsSearchConditional Then
                FindNearestAlternative(token, text)
                If token.HadError Then Return False
            End If

            ' when the first tag ends with the start of what follows, step past it so it is not reused
            If token.Role = TokenRole.Open Then
                Dim following = If(token.NextTagValue.Length = 0, token.Parent.CloseToken.FirstTagValue, token.NextTagValue)
                If following.Length > 0 AndAlso token.FirstTagValue.EndsWith(following, StringComparison.Ordinal) Then
                    _position += 1
                End If
            End If

            If token.HasJump Then
                Dim target = token.NextTagValue
                If target.Length > 0 AndAlso token.FirstTagValue.EndsWith(target(0)) Then
                    _position += 1
                End If
                token.NextTagPos = IndexOf(text, target, _position - 1)
                If token.HadError Then
                    If Debug Then Trace($"  {token}: jump target [{target}] not found")
                    Return False
                End If
                _position = If(token.IsLookahead, token.NextTagPos, token.NextPosEnd)
                Dim size = Math.Min(token.NextPrevCharsSize, token.NextPosEnd)
                token.NextTagPrevChars = text.Substring(token.NextPosEnd - size, size)
                If Debug Then Trace($"  {token}: jump [{target}] @{token.NextTagPos} prev chars [{token.NextTagPrevChars}]")
            End If
            Return True
        End Function

        ''' <summary>
        ''' For <c>%{}</c> tokens: find the next opening character (failing if the stop character
        ''' comes first) and the character that balances it. Nesting is counted; quoting is not.
        ''' </summary>
        Private Function MatchBalanced(token As ParserToken, text As String) As Boolean
            Dim openChar = token.BalancedOpen
            Dim closeChar = token.BalancedClose
            Dim stopChar = token.BalancedStop
            Dim i = _position
            While i < text.Length AndAlso text(i) <> openChar
                If stopChar <> ChrW(0) AndAlso text(i) = stopChar Then
                    token.HadError = True
                    Return False
                End If
                i += 1
            End While
            If i >= text.Length Then
                token.HadError = True
                Return False
            End If

            Dim j As Integer
            If openChar = closeChar Then
                j = text.IndexOf(closeChar, i + 1)
            Else
                Dim depth = 0
                j = i
                While j < text.Length
                    If text(j) = openChar Then
                        depth += 1
                    ElseIf text(j) = closeChar Then
                        depth -= 1
                        If depth = 0 Then Exit While
                    End If
                    j += 1
                End While
                If j >= text.Length Then j = -1
            End If
            If j < 0 Then
                token.HadError = True
                Return False
            End If

            token.SetBalancedMatch(i, j)
            _position = j + 1
            If Debug Then Trace($"  {token}: balanced {openChar}{closeChar} @{i}..{j}")
            Return True
        End Function

        ''' <summary>For <c>a|b|c</c> tokens: pick the nearest alternative that occurs after the first tag.</summary>
        Private Sub FindNearestAlternative(token As ParserToken, text As String)
            If text.Length = 0 Then
                token.NextTagPos = 0
                Return
            End If
            Dim bestPos = -1
            Dim bestValue = ""
            For Each alternative In token.NextTagConditions
                If alternative.Length = 0 Then Continue For
                Dim p = IndexOf(text, alternative, _position)
                If p < 0 Then Continue For
                Dim better = bestPos < 0 OrElse If(token.SeekBackward, p < bestPos, p > bestPos)
                If better Then
                    bestPos = p
                    bestValue = alternative
                End If
            Next
            ' a close token that finds none of its alternatives ends at the end of the text
            If bestPos < 0 AndAlso token.Role = TokenRole.Close Then bestPos = text.Length
            token.NextTagConditionValue = bestValue
            token.NextTagPos = bestPos
            If token.HadError Then Return
            token.NextPrevCharsSize = bestValue.Length
            _position = If(token.IsLookahead, bestPos, token.NextPosEnd)
            token.NextTagPrevChars = text.Substring(token.NextPosEnd - token.NextPrevCharsSize, token.NextPrevCharsSize)
            If Debug Then Trace($"  {token}: alternative [{bestValue}] @{bestPos}")
        End Sub

        ''' <summary>Finds the first tag, rejecting hits where the field runs into neighbouring name characters.</summary>
        Private Shared Function FindFirstTag(token As ParserToken, text As String, start As Integer) As Integer
            Dim value = token.FirstTagValue
            Dim trailing = token.FieldNeedsTrailingBoundary
            Dim leading = token.FieldNeedsLeadingBoundary
            Dim pos = IndexOf(text, value, start)
            While pos > -1
                Dim endPos = pos + value.Length
                Dim ok = Not (trailing AndAlso endPos < text.Length AndAlso IsNameChar(text(endPos))) AndAlso
                         Not (leading AndAlso pos > 0 AndAlso IsNameChar(text(pos - 1)))
                If ok Then Return pos
                pos = IndexOf(text, value, pos + 1)
            End While
            Return -1
        End Function

        ''' <summary>
        ''' For a token whose field is empty (or <c>*</c>), read the field name from the text.
        ''' The token's first tag is <c>lead*trail</c>: with a lead, the field is the run of name
        ''' characters right after the next lead (so <c>&lt;*</c> finds <c>&lt;variable</c> but skips
        ''' <c>&lt;/variable</c> and <c>&lt;!--</c>); with no lead, it is the run of name characters
        ''' right before the next trail (so <c>*="</c> finds <c>name="</c>).
        ''' Returns False when no further field can be found.
        ''' </summary>
        Private Function DeriveWildcardField(token As ParserToken, text As String) As Boolean
            Dim lead As String = Nothing, trail As String = Nothing
            token.GetFieldContext(lead, trail)
            If lead.Length = 0 AndAlso trail.Length = 0 Then Return False

            Dim searchFrom = Math.Max(0, _position)
            Do
                Dim fieldStart As Integer, fieldEnd As Integer, anchor As Integer
                If lead.Length > 0 Then
                    anchor = IndexOf(text, lead, searchFrom)
                    If anchor < 0 Then Return False
                    fieldStart = anchor + lead.Length
                    fieldEnd = fieldStart
                    While fieldEnd < text.Length AndAlso IsNameChar(text(fieldEnd))
                        fieldEnd += 1
                    End While
                Else
                    Dim trailPos = IndexOf(text, trail, searchFrom)
                    If trailPos < 0 Then Return False
                    fieldEnd = trailPos
                    fieldStart = trailPos
                    While fieldStart > searchFrom AndAlso IsNameChar(text(fieldStart - 1))
                        fieldStart -= 1
                    End While
                    anchor = fieldStart
                End If

                Dim followsOk = trail.Length = 0 OrElse
                    String.CompareOrdinal(text, fieldEnd, trail, 0, trail.Length) = 0
                If fieldEnd > fieldStart AndAlso followsOk Then
                    Dim field = text.Substring(fieldStart, fieldEnd - fieldStart)
                    token.Parent.OpenToken.FieldValue = field
                    token.Parent.CloseToken.FieldValue = field
                    token.Parent.IsWildCarding = True
                    _position = anchor
                    If Debug Then Trace($"  wildcard field derived: [{field}] @{anchor}")
                    Return True
                End If
                searchFrom = Math.Max(anchor, searchFrom) + 1
            Loop
        End Function

#End Region

#Region "Filters"

        ''' <summary>The parts of a filter such as <c>@type='number'</c>.</summary>
        Friend Structure FilterSpec
            Public Left As String
            Public Op As Char      ' = ! ~, X (exists), N (absent or empty)
            Public Right As String
        End Structure

        Friend Shared Function ParseFilter(filter As String) As FilterSpec
            For Each op In {"="c, "!"c, "~"c}
                Dim i = IndexOfUnquoted(filter, op)
                If i >= 1 Then
                    Dim right = filter.Substring(i + 1)
                    If op = "!"c AndAlso right.Length = 0 Then
                        Return New FilterSpec With {.Left = filter.Substring(0, i), .Op = "N"c, .Right = ""}
                    End If
                    Return New FilterSpec With {.Left = filter.Substring(0, i), .Op = op, .Right = right}
                End If
            Next
            If filter.StartsWith("!"c) AndAlso filter.Length > 1 Then
                Return New FilterSpec With {.Left = filter.Substring(1), .Op = "N"c, .Right = ""}
            End If
            Return New FilterSpec With {.Left = filter, .Op = "X"c, .Right = ""}
        End Function

        Private Shared Function IndexOfUnquoted(s As String, c As Char) As Integer
            Dim quoted = False
            For i = 0 To s.Length - 1
                If s(i) = "'"c Then
                    quoted = Not quoted
                ElseIf s(i) = c AndAlso Not quoted Then
                    Return i
                End If
            Next
            Return -1
        End Function

        ''' <summary>
        ''' Keeps the results whose left-hand path satisfies the filter. The left side is evaluated
        ''' inside each result; the right side is a quoted literal, a <c>//</c> path over the whole
        ''' text, or a relative path over the parent block. Comparisons ignore case and line breaks.
        ''' </summary>
        Private Function ApplyFilter(filter As String, parentText As String, results As List(Of Match)) As List(Of Match)
            Dim started = Stopwatch.GetTimestamp()
            Dim spec = ParseFilter(filter)

            Dim right As IReadOnlyList(Of String) = Array.Empty(Of String)()
            If spec.Op = "="c OrElse spec.Op = "!"c OrElse spec.Op = "~"c Then
                If spec.Right.Length >= 2 AndAlso spec.Right.StartsWith("'"c) AndAlso spec.Right.EndsWith("'"c) Then
                    ' a literal; a sub-path that produced several values arrives as one value per line
                    right = spec.Right.Substring(1, spec.Right.Length - 2).Split(ChrW(10)).Select(Function(v) v.Trim(ChrW(13))).ToList()
                Else
                    Dim scope = If(spec.Right.StartsWith("//", StringComparison.Ordinal), Text, " " & parentText & " ")
                    Dim rightParser = CreateChild(scope)
                    right = rightParser.Query(spec.Right)
                    Statistics.Add(rightParser.Statistics, includeSegments:=False)
                End If
            End If
            Dim rightNorm = right.Select(AddressOf NormalizeForCompare).ToList()
            If Verbose OrElse Debug Then Trace($" Filter: [{spec.Left}] {spec.Op} [{String.Join(" | ", right)}]")

            Dim kept = New List(Of Match)
            For Each result In results
                Dim leftParser = CreateChild(" " & result.Value & " ")
                Dim left = leftParser.Query(spec.Left).Select(AddressOf NormalizeForCompare).ToList()
                Statistics.Add(leftParser.Statistics, includeSegments:=False)

                Dim keep As Boolean
                Select Case spec.Op
                    Case "="c
                        keep = left.Any(Function(l) rightNorm.Contains(l))
                    Case "!"c
                        keep = Not left.Any(Function(l) rightNorm.Contains(l))
                    Case "~"c
                        keep = left.Any(Function(l) rightNorm.Any(Function(r) l.Contains(r, StringComparison.Ordinal)))
                    Case "X"c
                        keep = left.Count > 0
                    Case Else ' N
                        keep = left.Count = 0
                End Select
                If keep Then kept.Add(result)
            Next
            Statistics.FilterMilliseconds += Stopwatch.GetElapsedTime(started).TotalMilliseconds
            Return kept
        End Function

        Private Shared Function NormalizeForCompare(value As String) As String
            Return value.Replace(vbCr, "").Replace(vbLf, "").ToLowerInvariant()
        End Function

#End Region

#Region "Text helpers"

        Private Shared Function IndexOf(text As String, value As String, start As Integer) As Integer
            If start < 0 Then start = 0
            If start > text.Length Then Return -1
            Return text.IndexOf(value, start, StringComparison.Ordinal)
        End Function

        Private Shared Function IsBlank(c As Char) As Boolean
            Return c = " "c OrElse c = vbTab(0) OrElse c = vbLf(0) OrElse c = vbCr(0)
        End Function

        Friend Shared Function IsNameChar(c As Char) As Boolean
            Return Char.IsLetterOrDigit(c) OrElse c = "_"c OrElse c = "-"c OrElse c = "."c OrElse c = ":"c
        End Function

        Private Sub Trace(message As String)
            Log?.WriteLine(message)
        End Sub

#End Region

    End Class

End Namespace
