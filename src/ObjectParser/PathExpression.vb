Imports System.Text

Namespace ObjectParser

    ''' <summary>One <c>/</c>-separated step of a path, e.g. <c>@name</c> or <c>variable[flags]</c>.</summary>
    Public NotInheritable Class PathSegment
        Friend Sub New(text As String, prefix As String, field As String, filter As String)
            Me.Text = text
            Me.Prefix = prefix
            Me.Field = field
            Me.Filter = filter
        End Sub

        ''' <summary>The raw segment text, prefix included.</summary>
        Public ReadOnly Property Text As String
        ''' <summary>The token map prefix this segment uses (<c>" "</c> for the default pair).</summary>
        Public ReadOnly Property Prefix As String
        ''' <summary>The field name substituted for <c>*</c> in the tokens (may hold a <c>{:a,b:}</c> list).</summary>
        Public Property Field As String
        ''' <summary>The filter between <c>[</c> and <c>]</c>, without a leading <c>-</c>/<c>+</c>; empty when none.</summary>
        Public Property Filter As String

        Public ReadOnly Property HasFilter As Boolean
            Get
                Return Filter.Length > 0
            End Get
        End Property

        Public Overrides Function ToString() As String
            Return Text
        End Function
    End Class

    ''' <summary>
    ''' A parsed path such as <c>//variable[@type='number']/@name</c>.
    ''' </summary>
    ''' <remarks>
    ''' A leading <c>//</c> or <c>/</c> is accepted and ignored: every path searches the whole buffer.
    ''' One sub-path <c>{{...}}</c> may appear; it is evaluated first and its results replace it.
    ''' </remarks>
    Public NotInheritable Class PathExpression

        Private ReadOnly _syntax As ParserSyntax

        Private Sub New(syntax As ParserSyntax, original As String)
            _syntax = syntax
            Me.Original = original
        End Sub

        Public ReadOnly Property Original As String
        Public ReadOnly Property Segments As New List(Of PathSegment)
        ''' <summary>The <c>{{...}}</c> path, or empty.</summary>
        Public Property SubPath As String = ""

        ''' <summary>Parses <paramref name="path"/>, matching segment prefixes against <paramref name="prefixes"/> (longest first).</summary>
        Public Shared Function Parse(path As String, prefixes As IReadOnlyList(Of String), Optional syntax As ParserSyntax = Nothing) As PathExpression
            syntax = If(syntax, New ParserSyntax())
            Dim expr = New PathExpression(syntax, path)
            Dim body = If(path, "")
            If body.StartsWith("//", StringComparison.Ordinal) Then
                body = body.Substring(2)
            ElseIf body.StartsWith("/", StringComparison.Ordinal) Then
                body = body.Substring(1)
            End If

            ' lift out the sub-path so its separators and brackets do not split the main path
            Dim subOpen = body.IndexOf(syntax.SubPathOpen, StringComparison.Ordinal)
            If subOpen > -1 Then
                Dim subClose = body.IndexOf(syntax.SubPathClose, subOpen, StringComparison.Ordinal)
                If subClose > -1 Then
                    Dim start = subOpen + syntax.SubPathOpen.Length
                    expr.SubPath = body.Substring(start, subClose - start)
                    body = body.Substring(0, start) & ParserSyntax.SubPathReference & body.Substring(subClose)
                End If
            End If

            For Each part In SplitOutsideFilters(body, syntax)
                expr.Segments.Add(MakeSegment(part, prefixes, syntax))
            Next
            Return expr
        End Function

        ''' <summary>Splits on the path separator, ignoring separators inside <c>[...]</c> (nesting allowed).</summary>
        Private Shared Function SplitOutsideFilters(body As String, syntax As ParserSyntax) As List(Of String)
            Dim parts = New List(Of String)
            Dim current = New StringBuilder()
            Dim depth = 0
            Dim i = 0
            While i < body.Length
                If String.CompareOrdinal(body, i, syntax.FilterOpen, 0, syntax.FilterOpen.Length) = 0 Then
                    depth += 1
                ElseIf depth > 0 AndAlso String.CompareOrdinal(body, i, syntax.FilterClose, 0, syntax.FilterClose.Length) = 0 Then
                    depth -= 1
                ElseIf depth = 0 AndAlso String.CompareOrdinal(body, i, syntax.PathSeparator, 0, syntax.PathSeparator.Length) = 0 Then
                    parts.Add(current.ToString())
                    current.Clear()
                    i += syntax.PathSeparator.Length
                    Continue While
                End If
                current.Append(body(i))
                i += 1
            End While
            parts.Add(current.ToString())
            Return parts
        End Function

        Private Shared Function MakeSegment(text As String, prefixes As IReadOnlyList(Of String), syntax As ParserSyntax) As PathSegment
            Dim prefix = ParserSyntax.DefaultPrefix
            For Each p In prefixes
                If p.Length > 0 AndAlso text.StartsWith(p, StringComparison.Ordinal) Then
                    prefix = p
                    Exit For
                End If
            Next

            ' strip a non-default prefix, except for the current-result token and a bare field token
            Dim rest = text
            If prefix <> ParserSyntax.DefaultPrefix AndAlso
               Not text.StartsWith(syntax.CurrentResultToken, StringComparison.Ordinal) AndAlso
               text <> syntax.FieldToken Then
                rest = text.Substring(prefix.Length)
            End If

            Dim field = rest
            Dim filter = ""
            Dim open = rest.IndexOf(syntax.FilterOpen, StringComparison.Ordinal)
            If open > -1 AndAlso rest <> syntax.FieldToken Then
                field = rest.Substring(0, open)
                Dim close = rest.LastIndexOf(syntax.FilterClose, StringComparison.Ordinal)
                If close > open Then
                    filter = rest.Substring(open + syntax.FilterOpen.Length, close - open - syntax.FilterOpen.Length)
                    ' [+x] is the old "outer filter" marker and is ignored. A leading - is kept: it is
                    ' the inner-content prefix in the XML maps, so [-td='QA'] compares cell text.
                    If filter.StartsWith("+"c) Then
                        filter = filter.Substring(1)
                    End If
                End If
            End If
            Return New PathSegment(text, prefix, field, filter)
        End Function

        ''' <summary>Replaces the sub-path placeholder (and its braces) with <paramref name="value"/>.</summary>
        Friend Sub ApplySubPathResult(value As String)
            Dim placeholder = _syntax.SubPathOpen & ParserSyntax.SubPathReference & _syntax.SubPathClose
            For Each segment In Segments
                segment.Field = segment.Field.Replace(placeholder, value, StringComparison.Ordinal)
                segment.Filter = segment.Filter.Replace(placeholder, value, StringComparison.Ordinal)
            Next
        End Sub

        ''' <summary>The names in a <c>{:a,b:}</c> multi-path field, or Nothing when the field is not one.</summary>
        Friend Function MultiPathNames(segment As PathSegment, ByRef previous As String) As String()
            Dim f = segment.Field
            Dim open = f.IndexOf(_syntax.MultiPathOpen, StringComparison.Ordinal)
            Dim close = f.IndexOf(_syntax.MultiPathClose, StringComparison.Ordinal)
            If open < 0 OrElse close < 0 Then Return Nothing
            Dim start = open + _syntax.MultiPathOpen.Length
            Dim list = If(close > start, f.Substring(start, close - start), "")
            If list = _syntax.MultiPathPrevious Then
                list = previous
            ElseIf list.Length > 0 Then
                previous = list
            End If
            Return list.Split(_syntax.MultiPathDelimiter).Where(Function(n) n.Length > 0).ToArray()
        End Function

        Public Overrides Function ToString() As String
            Return Original
        End Function

    End Class

End Namespace
