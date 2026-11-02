Imports ObjectParser
Imports Xunit

''' <summary>How token strings are classified, and that each kind of token matches as documented.</summary>
Public Class TokenGrammarTests

    Private Shared Function MakePair(openToken As String, closeToken As String, Optional io As String = "I") As TokenPair
        Return TokenMap.Parse($"t={openToken}{vbLf}t={closeToken}{vbLf}t={io}").CreatePair("t")
    End Function

    <Fact>
    Public Sub Literal_token()
        Dim t = MakePair("<b>", "</b>").OpenToken
        Assert.False(t.IsComplex)
        Assert.False(t.HasJump)
        Assert.Equal("<b>", t.FirstTagToken)
        Assert.Equal("", t.NextTagToken)
    End Sub

    <Fact>
    Public Sub Jump_token_splits_on_caret()
        Dim t = MakePair("<*^>", "</*^>").OpenToken
        Assert.True(t.HasJump)
        Assert.Equal("<*", t.FirstTagToken)
        Assert.Equal(">", t.NextTagToken)
    End Sub

    <Fact>
    Public Sub Compare_conditional_token()
        Dim t = MakePair("<*^>", "/>?/>|</*^>").CloseToken
        Assert.True(t.HasCondition)
        Assert.True(t.IsCompareConditional)
        Assert.Equal("/>", t.FirstTagToken)
        Assert.Equal("/>|</*^>", t.NextTagToken)
    End Sub

    <Fact>
    Public Sub Search_conditional_token()
        Dim t = MakePair("<", " | |/|>").CloseToken
        Assert.True(t.IsSearchConditional)
        Assert.Equal(" ", t.FirstTagToken)
    End Sub

    <Fact>
    Public Sub Question_mark_without_alternatives_is_literal()
        Dim t = MakePair("<?", "?>").OpenToken
        Assert.False(t.HasCondition)
        Assert.Equal({"xml version=""1.0"""}, QMap("p=<?" & vbLf & "p=?>" & vbLf & "p=I", "<?xml version=""1.0""?>", "//p"))
    End Sub

    <Fact>
    Public Sub Balanced_token()
        Dim t = MakePair("class *", "%{};").CloseToken
        Assert.True(t.IsBalanced)
    End Sub

    <Fact>
    Public Sub Jump_skips_to_target()
        ' find "key=" then jump to the next quote; inner text runs from after the jump target
        Dim map = "k=key=^""" & vbLf & "k=""" & vbLf & "k=I"
        Assert.Equal({"abc", "x y"}, QMap(map, "key=  ""abc"" key=""x y""", "//k"))
    End Sub

    <Fact>
    Public Sub Compare_conditional_handles_self_closing_and_paired_tags()
        Dim xml = "<r><a id=""1""/><a id=""2"">two</a><a id=""3""/></r>"
        Assert.Equal({"<a id=""1""/>", "<a id=""2"">two</a>", "<a id=""3""/>"}, Q(xml, "//a"))
    End Sub

    <Fact>
    Public Sub Search_conditional_picks_nearest_alternative()
        Dim map = "w=:" & vbLf & "w=|;|,|." & vbLf & "w=I"
        Assert.Equal({"one", "two", "three"}, QMap(map, ":  one; : two,: three.", "//w").Select(Function(s) s.Trim()).ToList())
    End Sub

    <Fact>
    Public Sub Lookahead_close_does_not_consume_its_target()
        ' each item ends at the next "\n-" which also starts the following item
        Dim map = "i=\n-\s" & vbLf & "i=|\n|\n" & vbLf & "i=I"
        Assert.Equal({"a", "b", "c"}, QMap(map, vbLf & "- a" & vbLf & "- b" & vbLf & "- c", "//i"))
    End Sub

    <Fact>
    Public Sub Lookahead_close_runs_to_end_of_text_when_nothing_follows()
        Dim map = "c=//" & vbLf & "c=|\r|\n" & vbLf & "c=I"
        Assert.Equal({" one", " two"}, QMap(map, "x // one" & vbLf & "y // two", "//c"))
    End Sub

    <Fact>
    Public Sub Balanced_braces_count_nesting()
        Dim code = "class A { void F() { if (x) { y(); } } } class B { }"
        Assert.Equal({"class A { void F() { if (x) { y(); } } }"}, Q(code, "//A", "cs"))
        Assert.Equal({" void F() { if (x) { y(); } } "}, Q(code, "//-A", "cs"))
    End Sub

    <Fact>
    Public Sub Balanced_stop_character_rejects_candidates()
        ' "F();" is a call (a ; comes before any {); "F() {" is the declaration
        Dim code = "class A { void G() { F(); } void F() { z(); } }"
        Assert.Equal({"F() { z(); }"}, Q(code, "//A/!F", "cs"))
    End Sub

    <Fact>
    Public Sub Balanced_parentheses_and_same_char_pairs()
        Dim map = "p=call" & vbLf & "p=%()" & vbLf & "p=I" & vbLf & "q=say" & vbLf & "q=%''" & vbLf & "q=I"
        Assert.Equal({"a, (b, c)"}, QMap(map, "call(a, (b, c)) tail", "//p"))
        Assert.Equal({"hi there"}, QMap(map, "say 'hi there' done", "//q"))
    End Sub

    <Fact>
    Public Sub Field_scope_returns_the_name_read_from_text()
        Dim map = "n=<*" & vbLf & "n=|>|/" & vbLf & "n=F"
        Assert.Equal({"a", "b", "c"}, QMap(map, "<a><b/></a><!-- x --><c></c>", "//n"))
    End Sub

    <Fact>
    Public Sub Escaped_specials_match_literally()
        Dim map = "s=\*\*" & vbLf & "s=\*\*" & vbLf & "s=I" & vbLf & "p=a\|b" & vbLf & "p=\^" & vbLf & "p=I"
        Assert.Equal({"bold"}, QMap(map, "x **bold** y", "//s"))
        Assert.Equal({"-"}, QMap(map, "a|b-^", "//p"))
    End Sub

    <Fact>
    Public Sub Field_boundaries_prevent_partial_name_matches()
        Dim xml = "<r><choices><choice>A</choice></choices><x rename=""no"" name=""yes""/></r>"
        Assert.Equal({"<choice>A</choice>"}, Q(xml, "//choice"))
        Assert.Equal({"yes"}, Q(xml, "//x/@name"))
    End Sub

End Class
