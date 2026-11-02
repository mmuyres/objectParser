Imports ObjectParser
Imports Xunit

''' <summary>The original use case: pulling process data out of wiki pages, as HTML and as MediaWiki markup.</summary>
Public Class WikiAndHtmlTests

    Private Shared ReadOnly Html As String = ReadData("wiki.html")
    Private Shared ReadOnly Markup As String = ReadData("release.wiki")
    Private Shared ReadOnly WikiMap As TokenMap = TokenMap.FromFile(DataPath("mediawiki.map"))

    Private Shared Function H(path As String) As List(Of String)
        Return LfAll(Q(Html, path, "html"))
    End Function
    Private Shared Function W(path As String) As List(Of String)
        Return LfAll(New Parser(WikiMap, Markup).Query(path))
    End Function

    <Fact>
    Public Sub Html_list_items_and_classes()
        Assert.Equal({"Cut the branch", "Run the <a href=""/wiki/Tests"">tests</a>", "Tag the build"}, H("//-li"))
        Assert.Equal({"step", "step", "step done"}, H("//li/@class"))
        Assert.Equal({"<li class=""step done"">Tag the build</li>"}, H("//li[@class~'done']"))
    End Sub

    <Fact>
    Public Sub Html_links_headings_and_title()
        Assert.Equal({"/wiki/Tests", "/wiki/Rollback", "/wiki/Hotfix"}, H("//a/@href"))
        Assert.Equal({"Steps", "Owners"}, H("//-h2"))
        Assert.Equal({"Release Process"}, H("//-title"))
    End Sub

    <Fact>
    Public Sub Html_table_cells()
        Assert.Equal({"Build", "Pat", "QA", "Sam"}, H("//table/tr/-td"))
        Assert.Equal({"Role", "Name"}, H("//table/tr/-th"))
        ' rows that have a cell equal to "QA"
        Assert.Equal({"<tr><td>QA</td><td>Sam</td></tr>"}, H("//table/tr[-td='QA']"))
    End Sub

    <Fact>
    Public Sub Html_filter_by_class()
        Assert.Equal({"How we ship a <b>release</b>."}, H("//-p").Take(1))
        Assert.Equal({"<p class=""intro"">How we ship a <b>release</b>.</p>"}, H("//p[@class='intro']"))
    End Sub

    <Fact>
    Public Sub Wiki_headings_links_and_templates()
        Assert.Equal({"Steps", "Owners"}, W("//h2"))
        Assert.Equal({"Escalation"}, W("//h3"))
        Assert.Equal({"Rollback", "Hotfix", "Tests", "On-call"}, W("//target"))
        Assert.Equal({"Infobox process|owner=Build team|status=active"}, W("//tpl"))
        Assert.Equal({"Release Process"}, W("//b"))
    End Sub

    <Fact>
    Public Sub Wiki_sections_and_bullets()
        Assert.Equal({"Cut the branch", "Run the [[Tests|test suite]]", "Tag the build", "Build: Pat", "QA: Sam"}, W("//li"))
        Assert.Equal({"Build: Pat", "QA: Sam"}, W("//Owners/li"))
        Assert.Equal({"Tests|test suite"}, W("//Steps/link"))
        Assert.Equal({"QA: Sam"}, W("//li[.~'qa']"))
    End Sub

End Class
