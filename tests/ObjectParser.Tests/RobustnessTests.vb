Imports ObjectParser
Imports Xunit

''' <summary>
''' Random and malformed input must never hang, throw, or report a match that is not really
''' at the position it claims.
''' </summary>
Public Class RobustnessTests

    Private Shared ReadOnly Queries As New Dictionary(Of String, String()) From {
        {"xml", {"//a", "//-a", "//a/@id", "//a/b", "//$", "//#", "//!", "//a[@id='x']", "//a[b!]", "//{:a,b:}/@id", "//", "//*", "//a/"}},
        {"xml_ext", {"//comment:", "//node:a", "//nodename:", "//a/", "//cdata:"}},
        {"vb", {"//:", "//!", "//@", "//X/p", "//v", "//#", "//m"}},
        {"cs", {"//:", "//m", "//!", "//X/!F/(", "//c", "//b", "//u"}}
    }

    Private Shared Function RandomText(rnd As Random, alphabet As String, length As Integer) As String
        Return New String(Enumerable.Range(0, length).Select(Function(i) alphabet(rnd.Next(alphabet.Length))).ToArray())
    End Function

    Public Shared ReadOnly Property Seeds As IEnumerable(Of Object())
        Get
            Return Enumerable.Range(1, 60).Select(Function(s) New Object() {s})
        End Get
    End Property

    <Theory>
    <MemberData(NameOf(Seeds))>
    Public Sub Random_input_terminates_and_matches_are_genuine(seed As Integer)
        Dim rnd = New Random(seed)
        Dim alphabets = {"<>/=""' abid" & vbLf & "!-?[]", "{}();/* abX" & vbLf & "classF", "ClassFunctionEndSub As(). XabByVal " & vbLf}
        Dim work = Task.Run(
            Sub()
                For Each alphabet In alphabets
                    Dim text = RandomText(rnd, alphabet, rnd.Next(0, 400))
                    For Each kv In Queries
                        Dim parser = New Parser(TokenMap.BuiltIn(kv.Key), text)
                        For Each query In kv.Value
                            For Each match In parser.QueryMatches(query)
                                Assert.InRange(match.Start, 0, text.Length)
                                Assert.Equal(match.Value, text.Substring(match.Start, match.Length))
                            Next
                        Next
                    Next
                Next
            End Sub)
        Assert.True(work.Wait(TimeSpan.FromSeconds(20)), "parser did not finish (possible infinite loop)")
    End Sub

    <Theory>
    <InlineData("")>
    <InlineData("<")>
    <InlineData("<a")>
    <InlineData("<a>")>
    <InlineData("</a>")>
    <InlineData("<a/")>
    <InlineData("<a id=""")>
    <InlineData("<a><a><a>")>
    <InlineData(">>>><<<<")>
    Public Sub Truncated_xml_is_handled(text As String)
        Dim parser = New Parser(TokenMap.BuiltIn("xml"), text)
        For Each query In Queries("xml")
            Dim results = parser.QueryMatches(query)
            Assert.All(results, Sub(match) Assert.Equal(match.Value, text.Substring(match.Start, match.Length)))
        Next
    End Sub

    <Theory>
    <InlineData("//a[")>
    <InlineData("//a]")>
    <InlineData("//a[@id=")>
    <InlineData("//{{")>
    <InlineData("//{{//a")>
    <InlineData("//{:a,")>
    <InlineData("////")>
    <InlineData("//a[']")>
    <InlineData("")>
    Public Sub Odd_paths_do_not_throw(path As String)
        Dim parser = New Parser(TokenMap.BuiltIn("xml"), "<a id=""1""><b/></a>")
        parser.Query(path)
    End Sub

    <Fact>
    Public Sub Large_document_scales_linearly_enough()
        Dim sb = New Text.StringBuilder("<root>")
        For i = 1 To 20000
            sb.Append($"<item id=""{i}""><name>n{i}</name></item>")
        Next
        sb.Append("</root>")
        Dim parser = New Parser(TokenMap.BuiltIn("xml"), sb.ToString())
        Dim sw = Stopwatch.StartNew()
        Dim ids = parser.Query("//item/@id")
        Dim names = parser.Query("//item/-name")
        sw.Stop()
        Assert.Equal(20000, ids.Count)
        Assert.Equal("20000", ids.Last())
        Assert.Equal("n20000", names.Last())
        Assert.True(sw.Elapsed < TimeSpan.FromSeconds(10), $"took {sw.Elapsed}")
    End Sub

    <Fact>
    Public Sub Parser_is_reusable_across_queries()
        Dim parser = New Parser(TokenMap.BuiltIn("xml"), "<a id=""1""/><a id=""2""/>")
        For i = 1 To 5
            Assert.Equal({"1", "2"}, parser.Query("//a/@id"))
            Assert.Equal(2, parser.Query("//a").Count)
        Next
    End Sub

End Class
