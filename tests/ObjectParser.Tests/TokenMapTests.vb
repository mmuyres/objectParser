Imports ObjectParser
Imports Xunit

Public Class TokenMapTests

    <Fact>
    Public Sub Parses_three_line_entries_per_prefix()
        Dim map = TokenMap.Parse("a=<<" & vbLf & "a=>>" & vbLf & "a=I" & vbLf &
                                 "b=[" & vbCrLf & "b=]" & vbCrLf & "b=O" & vbCrLf)
        Assert.True(map.Contains("a"))
        Assert.True(map.Contains("b"))
        Assert.Empty(map.Warnings)
        Dim b = map.Definitions.Single(Function(d) d.Prefix = "b")
        Assert.Equal("[", b.OpenToken)
        Assert.Equal("]", b.CloseToken)
        Assert.Equal("O", b.IOSpec)
    End Sub

    <Fact>
    Public Sub Space_is_the_default_prefix_and_blank_prefix_maps_to_it()
        Dim map = TokenMap.BuiltIn("xml")
        Assert.True(map.Contains(" "))
        Assert.True(map.Contains(""))
        Assert.True(map.Contains(Nothing))
    End Sub

    <Fact>
    Public Sub Missing_io_line_defaults_to_inner()
        Dim map = TokenMap.Parse("q=<" & vbLf & "q=>")
        Assert.Equal("I", map.Definitions.Single(Function(d) d.Prefix = "q").IOSpec)
        Assert.Equal({"x"}, QMap("q=<" & vbLf & "q=>", "<x>", "//q"))
    End Sub

    <Fact>
    Public Sub Bad_entries_are_reported_not_thrown()
        Dim map = TokenMap.Parse("no equals sign here" & vbLf & "a=<" & vbLf & "b=>" & vbLf & "b=I")
        Assert.NotEmpty(map.Warnings)
        Assert.False(map.Contains("a"))
    End Sub

    <Fact>
    Public Sub Current_result_prefix_is_reserved()
        Dim map = TokenMap.Parse(".=<" & vbLf & ".=>" & vbLf & ".=I")
        Assert.Contains(map.Warnings, Function(w) w.Contains("Reserved"))
        Assert.Equal("O", map.Definitions.Single(Function(d) d.Prefix = ".").IOSpec)
    End Sub

    <Fact>
    Public Sub Field_token_as_prefix_disables_wildcards()
        Dim map = TokenMap.Parse("*=</" & vbLf & "*=>" & vbLf & "*=I")
        Assert.False(map.WildCardSupportEnabled)
        Assert.Contains(map.Warnings, Function(w) w.Contains("field token"))
        For Each name In BuiltInMaps.Names
            Assert.True(TokenMap.BuiltIn(name).WildCardSupportEnabled, name)
        Next
    End Sub

    <Fact>
    Public Sub Prefixes_are_ordered_longest_first()
        Dim map = TokenMap.BuiltIn("xml_ext")
        Dim lengths = map.Prefixes.Select(Function(p) p.Length).ToList()
        Assert.Equal(lengths.OrderByDescending(Function(l) l).ToList(), lengths)
        Assert.Equal("nodename:", map.Prefixes(0))
    End Sub

    <Fact>
    Public Sub Later_entries_replace_earlier_ones()
        Dim map = TokenMap.Parse("a=<" & vbLf & "a=>" & vbLf & "a=I" & vbLf & "a=[" & vbLf & "a=]" & vbLf & "a=I")
        Assert.Equal("[", map.Definitions.Single(Function(d) d.Prefix = "a").OpenToken)
    End Sub

    <Fact>
    Public Sub ToString_round_trips()
        Dim original = TokenMap.BuiltIn("vb")
        Dim copy = TokenMap.Parse(original.ToString())
        Assert.Equal(original.Definitions.Count, copy.Definitions.Count)
        For Each d In original.Definitions
            Dim c = copy.Definitions.Single(Function(x) x.Prefix = d.Prefix)
            Assert.Equal(d.OpenToken, c.OpenToken)
            Assert.Equal(d.CloseToken, c.CloseToken)
        Next
    End Sub

    <Theory>
    <InlineData("xml")>
    <InlineData("XML")>
    <InlineData("xml_ext")>
    <InlineData("html")>
    <InlineData("xhtml")>
    <InlineData("vb")>
    Public Sub Built_in_maps_load_cleanly(name As String)
        Dim map = TokenMap.BuiltIn(name)
        Assert.True(map.Definitions.Count > 2)
        Assert.Empty(map.Warnings)
    End Sub

    <Fact>
    Public Sub Unknown_built_in_map_throws_with_known_names()
        Dim ex = Assert.Throws(Of ArgumentException)(Function() TokenMap.BuiltIn("json"))
        Assert.Contains("xml_ext", ex.Message)
    End Sub

    <Fact>
    Public Sub Map_files_load_from_disk()
        Dim map = TokenMap.FromFile(DataPath("mediawiki.map"))
        Assert.Empty(map.Warnings)
        Assert.True(map.Contains("link"))
    End Sub

    <Fact>
    Public Sub Escapes_allow_line_breaks_and_tabs_in_tokens()
        Dim map = TokenMap.Parse("k=*: " & vbLf & "k=\n" & vbLf & "k=I")
        Dim d = map.Definitions.Single(Function(x) x.Prefix = "k")
        Assert.Equal(vbLf, d.CloseToken)
        Dim m2 = TokenMap.Parse("t=\t" & vbLf & "t=\\" & vbLf & "t=I")
        Dim t = m2.Definitions.Single(Function(x) x.Prefix = "t")
        Assert.Equal(vbTab, t.OpenToken)
        Assert.Equal("\", t.CloseToken)
    End Sub

End Class
