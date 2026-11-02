Imports ObjectParser
Imports Xunit

''' <summary>
''' The documented path examples, run against TestData/model.xml. Expected values were checked
''' against the original v4.1 executable; differences are the fixes listed in docs/changes.md.
''' </summary>
Public Class XmlQueryTests

    Private Shared ReadOnly Model As String = ReadData("model.xml")

    Private Shared Function M(path As String) As List(Of String)
        Return LfAll(Q(Model, path))
    End Function

    <Fact>
    Public Sub Attribute_values_by_name()
        Assert.Equal({"alpha", "beta", "gamma"}, M("//variable/@name"))
    End Sub

    <Fact>
    Public Sub Attribute_search_covers_the_whole_block_including_descendants()
        Assert.Equal({"pulldown", "checkbox"}, M("//model/controls/type/@value"))
        Assert.Equal({"demo", "idxA", "idxB", "alpha", "beta", "gamma", "d1", "o1", "a1", "CNS-989", "a2", "a3"}, M("//model/@name"))
    End Sub

    <Fact>
    Public Sub Every_attribute_value()
        Assert.Equal({"alpha", "number", "textnode", "beta", "beta", "text", "idxA", "gamma", "number"}, M("//variable/@"))
    End Sub

    <Fact>
    Public Sub Attribute_names()
        Assert.Equal({"name", "type", "attr", "value"}, M("//variable/$").Take(4))
        Assert.All(M("//variable/$"), Sub(n) Assert.Matches("^[a-z_]+$", n))
    End Sub

    <Fact>
    Public Sub Element_names()
        Assert.Equal({"flags", "flag", "node", "flags", "flag"}, M("//variable/#"))
    End Sub

    <Fact>
    Public Sub Whole_elements_in_document_order()
        Dim v = M("//variable")
        Assert.Equal(3, v.Count)
        Assert.StartsWith("<variable name=""alpha""", v(0))
        Assert.EndsWith("</variable>", v(0))
        Assert.StartsWith("<variable name=""beta""", v(1))
        Assert.Equal("<variable name=""gamma"" type=""number""/>", v(2))
    End Sub

    <Fact>
    Public Sub Inner_text()
        Assert.Equal({"first"}, M("//variable/-node"))
        Assert.Equal({"Pick", "Radio"}, M("//-choice"))
        Assert.Equal({"Pick", "Radio"}, M("//choice/!"))
    End Sub

    <Fact>
    Public Sub Element_name_does_not_match_longer_names()
        ' v4.1 returned "<choices>..." here because "<choice" matched "<choices"
        Assert.Equal({"<choice type=""pulldown"">Pick</choice>", "<choice type=""radio"">Radio</choice>"}, M("//choice"))
    End Sub

    <Fact>
    Public Sub Wildcard_children()
        Assert.Equal({"<att name=""a1"" type=""pulldown"" snapshot_output=""ind_one""><option id=""7"" name=""CNS-989""/></att>",
                      "<att name=""a2"" snapshot_output=""x""/>",
                      "<att name=""a3""/>"}, M("//projatt/"))
        Assert.Equal({"model"}, M("//*").Select(Function(e) e.Substring(1, 5)))
    End Sub

    <Fact>
    Public Sub Nested_paths()
        Assert.Equal({"portfolio_index", "other", "textnode", "idxA", "idxA", "apfile"}, M("//flags/flag/@attr"))
        Assert.Equal({"<node>first</node>"}, M("//node"))
    End Sub

    <Fact>
    Public Sub Multi_path_lists_several_names()
        Assert.Equal({"alpha", "beta", "gamma", "d1", "o1"}, M("//{:variable,decision,objective:}/@name"))
    End Sub

    <Fact>
    Public Sub Multi_path_previous_list_is_reusable()
        ' {:*:} re-uses the last list: inside each variable/decision, look for the same names again (none nest)
        Assert.Empty(M("//{:variable,decision:}/{:*:}"))
    End Sub

    <Fact>
    Public Sub Missing_elements_return_nothing()
        Assert.Empty(M("//nothing"))
        Assert.Empty(M("//variable/@missing"))
    End Sub

    <Fact>
    Public Sub Current_result_token_bridges_two_filters()
        Assert.Equal({"<option id=""7"" name=""CNS-989""/>"}, M("//projatt/att/option[@id='7']/.[@name='CNS-989']"))
    End Sub

    <Fact>
    Public Sub Results_token_can_be_changed()
        Dim syntax = New ParserSyntax() With {.CurrentResultToken = "~"}
        Dim parser = New Parser(TokenMap.BuiltIn("xml", syntax), Model)
        Assert.Equal({"a1", "CNS-989"}, parser.Query("//projatt/att[option]/~[@type='pulldown']/@name"))
    End Sub

    <Fact>
    Public Sub Xml_ext_named_prefixes()
        Dim ext = Function(p As String) LfAll(Q(Model, p, "xml_ext"))
        Assert.Equal({" a comment "}, ext("//comment:").Select(Function(c) c.Replace("<!--", "").Replace("-->", "")))
        Assert.Equal({"<variable name=""alpha"" type=""number"">", "<variable name=""beta"" type=""text"">", "<variable name=""gamma"" type=""number""/>"}, ext("//node:variable"))
        Assert.Equal({"name", "type"}, ext("//node:variable/attname:").Take(2))
        Assert.Equal({"version=""1.0"""}, ext("//pi:").Select(Function(s) s.Replace("xml ", "")))
    End Sub

    <Fact>
    Public Sub Cdata_content()
        Dim xml = "<a><![CDATA[x < y && z]]></a>"
        Assert.Equal({"x < y && z"}, Q(xml, "//cdata:", "xml_ext"))
    End Sub

    <Fact>
    Public Sub Empty_elements_do_not_stop_the_scan()
        ' v4.1 stopped at the first zero-length inner result
        Assert.Equal({"", "x", ""}, Q("<r><td></td><td>x</td><td></td></r>", "//-td"))
    End Sub

    <Fact>
    Public Sub Matches_report_exact_positions()
        Dim parser = New Parser(TokenMap.BuiltIn("xml"), Model)
        Dim matches = parser.QueryMatches("//variable/@type")
        Assert.Equal(3, matches.Count)
        For Each match In matches
            Assert.Equal(match.Value, Model.Substring(match.Start, match.Length))
        Next
    End Sub

End Class
