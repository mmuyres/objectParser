Imports ObjectParser
Imports Xunit

Public Class PathExpressionTests

    Private Shared ReadOnly XmlPrefixes As IReadOnlyList(Of String) = TokenMap.BuiltIn("xml").Prefixes

    <Fact>
    Public Sub Splits_segments_and_strips_prefixes()
        Dim p = PathExpression.Parse("//model/variable/@name", XmlPrefixes)
        Assert.Equal({"model", "variable", "name"}, p.Segments.Select(Function(s) s.Field))
        Assert.Equal({" ", " ", "@"}, p.Segments.Select(Function(s) s.Prefix))
    End Sub

    <Theory>
    <InlineData("//a/b")>
    <InlineData("/a/b")>
    <InlineData("a/b")>
    Public Sub Leading_slashes_are_optional(path As String)
        Assert.Equal({"a", "b"}, PathExpression.Parse(path, XmlPrefixes).Segments.Select(Function(s) s.Field))
    End Sub

    <Fact>
    Public Sub Separators_inside_filters_do_not_split()
        Dim p = PathExpression.Parse("//variable[flags/flag/@attr='x/y']/@name", XmlPrefixes)
        Assert.Equal(2, p.Segments.Count)
        Assert.Equal("variable", p.Segments(0).Field)
        Assert.Equal("flags/flag/@attr='x/y'", p.Segments(0).Filter)
    End Sub

    <Fact>
    Public Sub Nested_brackets_in_filters_are_balanced()
        Dim p = PathExpression.Parse("//a[b[c='1']/d]/e", XmlPrefixes)
        Assert.Equal({"a", "e"}, p.Segments.Select(Function(s) s.Field))
        Assert.Equal("b[c='1']/d", p.Segments(0).Filter)
    End Sub

    <Theory>
    <InlineData("//a[-b='1']", "-b='1'")>
    <InlineData("//a[+b='1']", "b='1'")>
    <InlineData("//a[b]", "b")>
    Public Sub Outer_marker_is_removed_and_minus_is_a_prefix(path As String, filter As String)
        Assert.Equal(filter, PathExpression.Parse(path, XmlPrefixes).Segments(0).Filter)
    End Sub

    <Fact>
    Public Sub Current_result_segment_keeps_its_token()
        Dim p = PathExpression.Parse("//a[x]/.[y]", XmlPrefixes)
        Assert.Equal(".", p.Segments(1).Prefix)
        Assert.Equal("y", p.Segments(1).Filter)
    End Sub

    <Fact>
    Public Sub Sub_path_is_lifted_out()
        Dim p = PathExpression.Parse("//v[@name='{{//index[flags/flag]/@name}}']/@x", XmlPrefixes)
        Assert.Equal("//index[flags/flag]/@name", p.SubPath)
        Assert.Equal(2, p.Segments.Count)
        Assert.Contains("--SUB-PATH--", p.Segments(0).Filter)
    End Sub

    <Fact>
    Public Sub Longest_prefix_wins()
        Dim prefixes = TokenMap.BuiltIn("xml_ext").Prefixes
        Dim p = PathExpression.Parse("//nodename:/node:x/-y", prefixes)
        Assert.Equal({"nodename:", "node:", "-"}, p.Segments.Select(Function(s) s.Prefix))
        Assert.Equal({"", "x", "y"}, p.Segments.Select(Function(s) s.Field))
    End Sub

    <Fact>
    Public Sub Filter_parsing_covers_every_operator()
        Dim eq = Parser.ParseFilter("@type='number'")
        Assert.Equal("="c, eq.Op)
        Assert.Equal("@type", eq.Left)
        Assert.Equal("'number'", eq.Right)

        Assert.Equal("!"c, Parser.ParseFilter("@type!'x'").Op)
        Assert.Equal("~"c, Parser.ParseFilter("@name~'ph'").Op)
        Assert.Equal("X"c, Parser.ParseFilter("flags/flag").Op)
        Assert.Equal("N"c, Parser.ParseFilter("option!").Op)

        Dim negated = Parser.ParseFilter("!flags")
        Assert.Equal("N"c, negated.Op)
        Assert.Equal("flags", negated.Left)
    End Sub

    <Fact>
    Public Sub Operators_inside_quotes_are_ignored()
        Dim f = Parser.ParseFilter("@name~'a=b'")
        Assert.Equal("~"c, f.Op)
        Assert.Equal("'a=b'", f.Right)
    End Sub

    <Fact>
    Public Sub Equals_takes_precedence_so_bang_prefix_on_left_works()
        Dim f = Parser.ParseFilter("flags/!='x'")
        Assert.Equal("="c, f.Op)
        Assert.Equal("flags/!", f.Left)
    End Sub

End Class
