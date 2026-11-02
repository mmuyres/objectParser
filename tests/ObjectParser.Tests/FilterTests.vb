Imports ObjectParser
Imports Xunit

''' <summary>Filter semantics, as documented in docs/path-syntax.md.</summary>
Public Class FilterTests

    Private Shared ReadOnly Model As String = ReadData("model.xml")

    Private Shared Function M(path As String) As List(Of String)
        Return LfAll(Q(Model, path))
    End Function

    <Fact>
    Public Sub Equals_literal()
        Assert.Equal({"alpha", "gamma"}, M("//variable[@type='number']/@name"))
    End Sub

    <Fact>
    Public Sub Comparisons_ignore_case()
        Assert.Equal({"alpha", "gamma"}, M("//variable[@type='NUMBER']/@name"))
    End Sub

    <Fact>
    Public Sub Not_equal_literal()
        Assert.Equal({"beta"}, M("//variable[@type!'number']/@name"))
    End Sub

    <Fact>
    Public Sub Contains_literal()
        Assert.Equal({"alpha"}, M("//variable[@name~'ph']/@name"))
    End Sub

    <Fact>
    Public Sub Exists()
        Assert.Equal({"alpha", "beta"}, M("//variable[flags/flag]/@name"))
        Assert.Equal({"a2"}, M("//projatt/att[option!]/.[@snapshot_output]/@name"))
    End Sub

    <Fact>
    Public Sub Absent_with_trailing_bang()
        Assert.Equal({"a3"}, M("//projatt/att[option!]/.[@snapshot_output!]/@name"))
    End Sub

    <Fact>
    Public Sub Absent_with_leading_bang()
        ' documented in v4.1's usage text but evaluated as an existence check there
        Assert.Equal({"gamma"}, M("//variable[!flags]/@name"))
    End Sub

    <Fact>
    Public Sub Right_side_global_path()
        Assert.Equal({"<choice type=""pulldown"">Pick</choice>"}, M("//choices/choice[@type=//controls/type/@value]"))
    End Sub

    <Fact>
    Public Sub Right_side_global_path_with_single_value()
        ' v4.1 threw a NullReferenceException when the right side produced exactly one value
        Assert.Equal({"<choice type=""pulldown"">Pick</choice>"}, M("//choices/choice[@type=//controls/type[@value='pulldown']/@value]"))
    End Sub

    <Fact>
    Public Sub Right_side_relative_path_uses_parent_block()
        ' variables whose flag value names a variable... evaluated inside each parent block
        Dim xml = "<g><ref to=""b""/><item name=""a""/><item name=""b""/></g><g><ref to=""a""/><item name=""a""/></g>"
        Assert.Equal({"b", "a"}, Q(xml, "//g/item[@name=ref/@to]/@name"))
    End Sub

    <Fact>
    Public Sub Not_equal_against_a_list_means_none_match()
        Assert.Equal({"<choice type=""radio"">Radio</choice>"}, M("//choices/choice[@type!//controls/type/@value]"))
    End Sub

    <Fact>
    Public Sub Contains_against_a_list_means_any_match()
        Dim xml = "<r><t v=""north-east""/><t v=""south""/><t v=""west""/><k w=""north""/><k w=""west""/></r>"
        Assert.Equal({"north-east", "west"}, Q(xml, "//t[@v~//k/@w]/@v"))
    End Sub

    <Fact>
    Public Sub Inner_filter_marker()
        Assert.Equal({"<flags><flag attr=""apfile""/></flags>"}, M("//objective/flags[-flag/@attr='apfile']"))
    End Sub

    <Fact>
    Public Sub Sub_path_value_feeds_the_filter()
        Assert.Equal({"beta", "d1"}, M("//{:variable,decision,objective:}[flags/flag/@attr='{{//index[flags/flag/@attr~'portfolio']/@name}}']/@name"))
    End Sub

    <Fact>
    Public Sub Sub_path_with_several_values_acts_as_a_list()
        ' both index names: idxA (used by beta and d1) and idxB (unused)
        Assert.Equal({"beta", "d1"}, M("//{:variable,decision:}[flags/flag/@attr='{{//index/@name}}']/@name"))
    End Sub

    <Fact>
    Public Sub Filters_on_several_segments()
        Assert.Equal({"a1", "CNS-989"}, M("//projatt/att[option]/.[@type='pulldown']/@name"))
    End Sub

    <Fact>
    Public Sub Literal_may_contain_operator_characters()
        Dim xml = "<r><e k=""a=b""/><e k=""a!b""/></r>"
        Assert.Equal({"a=b"}, Q(xml, "//e[@k='a=b']/@k"))
        Assert.Equal({"a!b"}, Q(xml, "//e[@k~'!']/@k"))
    End Sub

End Class
