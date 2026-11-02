Imports ObjectParser
Imports Xunit

''' <summary>The same engine on formats it has no built-in map for: describe the format, then query it.</summary>
Public Class CustomMapTests

    <Fact>
    Public Sub Markdown_links()
        Dim map = "## markdown links: [text](url)
text=[
text=]
text=I
url=](
url=)
url=I
link=[
link=)
link=O"
        Dim md = "See [the docs](https://example.com/docs) and [source](https://example.com/src)."
        Assert.Equal({"the docs", "source"}, QMap(map, md, "//text"))
        Assert.Equal({"https://example.com/docs", "https://example.com/src"}, QMap(map, md, "//url"))
        Assert.Equal({"[source](https://example.com/src)"}, QMap(map, md, "//link[url~'src']"))
    End Sub

    <Fact>
    Public Sub Key_value_config_with_sections()
        Dim map = "## INI-style config
## Name   a [section] body, up to the next [ or the end
 =\n[*]
 =|\n[|\n[
 =I
## @key   value of key = ... on its own line
@=\n*=
@=|\r|\n
@=I"
        Dim ini = vbLf & "[db]" & vbLf & "host=localhost" & vbLf & "port=5432" & vbLf & "[web]" & vbLf & "host=0.0.0.0" & vbLf & "port=8080"
        Assert.Equal({"localhost"}, QMap(map, ini, "//db/@host"))
        Assert.Equal({"8080"}, QMap(map, ini, "//web/@port"))
        Assert.Equal({"5432", "8080"}, QMap(map, ini, "//@port"))
        Assert.Equal({"0.0.0.0"}, QMap(map, ini, "//@host[.~'0.0']").Select(Function(s) s.Trim()))
    End Sub

    <Fact>
    Public Sub Log_lines_by_level()
        Dim map = "## level-tagged log lines: [LEVEL] message
## Name   message text of lines tagged [Name]
 =[*]\s
 =|\r|\n
 =I"
        Dim log = "[INFO] started" & vbLf & "[WARN] disk 91%" & vbLf & "[INFO] ready" & vbLf & "[ERROR] lost connection"
        Assert.Equal({"started", "ready"}, QMap(map, log, "//INFO"))
        Assert.Equal({"lost connection"}, QMap(map, log, "//ERROR"))
        Assert.Equal({"disk 91%"}, QMap(map, log, "//{:WARN,ERROR:}[.~'disk']"))
    End Sub

    <Fact>
    Public Sub Json_like_objects()
        Dim map = "## flat JSON-ish: ""key"": ""value"" and ""key"": {object}
@=""*"":\s""
@=""
@=I
 =""*"":\s
 =%{}
 =O"
        Dim json = "{ ""user"": { ""name"": ""Pat"", ""role"": ""build"" }, ""team"": { ""name"": ""Release"" } }"
        Assert.Equal({"Pat", "Release"}, QMap(map, json, "//@name"))
        Assert.Equal({"Pat"}, QMap(map, json, "//user/@name"))
        Assert.Equal({"build"}, QMap(map, json, "//user/@role"))
    End Sub

    <Fact>
    Public Sub Same_text_two_maps()
        ' one buffer, two views: html structure and a custom "data-" attribute map
        Dim html = "<div data-id=""7"" data-owner=""pat"">x</div><div data-id=""8"">y</div>"
        Assert.Equal({"7", "8"}, Q(html, "//div/@data-id"))
        Dim dataMap = "d=data-*=""" & vbLf & "d=""" & vbLf & "d=I"
        Assert.Equal({"7", "pat", "8"}, QMap(dataMap, html, "//d"))
    End Sub

End Class
