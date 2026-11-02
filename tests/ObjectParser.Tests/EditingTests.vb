Imports ObjectParser
Imports Xunit

''' <summary>Delete, replace, keep and clean operate on exact match positions.</summary>
Public Class EditingTests

    Private Const Doc As String = "<r><a n=""1"">x</a><b>a</b><a n=""2""/></r>"

    Private Shared Function NewParser(Optional text As String = Doc) As Parser
        Return New Parser(TokenMap.BuiltIn("xml"), text)
    End Function

    <Fact>
    Public Sub Delete_removes_each_match()
        Assert.Equal("<r><b>a</b></r>", NewParser().Delete("//a"))
    End Sub

    <Fact>
    Public Sub Delete_only_touches_matched_spans()
        ' v4.1 used String.Replace on the matched text, so deleting the value "1" would strip every "1"
        Dim p = NewParser("<r><a n=""1"">1</a></r>")
        Assert.Equal("<r><a n="""">1</a></r>", p.Delete("//a/@n"))
    End Sub

    <Fact>
    Public Sub Replace_puts_text_in_place_of_each_match()
        Assert.Equal("<r>[A]<b>a</b>[A]</r>", NewParser().Delete("//a", replacement:="[A]"))
    End Sub

    <Fact>
    Public Sub Keep_replaces_each_match_with_a_sub_result()
        Assert.Equal("<r>1<b>a</b>2</r>", NewParser().Delete("//a", keepPath:="@n"))
    End Sub

    <Fact>
    Public Sub Nested_matches_are_removed_once()
        Dim p = NewParser("<r><x><x2/></x><x>t</x></r>")
        Assert.Equal("<r></r>", p.Delete("//x"))
    End Sub

    <Fact>
    Public Sub Delete_updates_text_for_later_queries()
        Dim p = NewParser()
        p.Delete("//b")
        Assert.Equal({"1", "2"}, p.Query("//a/@n"))
        Assert.Empty(p.Query("//b"))
    End Sub

    <Fact>
    Public Sub Clean_removes_noise_before_querying()
        Dim xml = "<r><!-- <a n=""ghost""/> --><a n=""real""/></r>"
        Dim p = New Parser(TokenMap.BuiltIn("xml_ext"), xml)
        Assert.Equal({"ghost", "real"}, p.Query("//a/@n"))
        p.Clean({"//comment:"})
        Assert.Equal({"real"}, p.Query("//a/@n"))
    End Sub

    <Fact>
    Public Sub Delete_with_no_matches_leaves_text_unchanged()
        Assert.Equal(Doc, NewParser().Delete("//zzz"))
    End Sub

    <Fact>
    Public Sub Strip_comments_from_csharp()
        Dim code = "int a; // note" & vbLf & "/* block */int b;"
        Dim p = New Parser(TokenMap.BuiltIn("cs"), code)
        p.Delete("//c")
        p.Delete("//b")
        Assert.Equal("int a; //" & vbLf & "/**/int b;", p.Text)
    End Sub

End Class
