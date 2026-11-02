Imports System.IO
Imports ObjParserCli
Imports Xunit

Public Class CliTests

    Private Structure RunResult
        Public Code As Integer
        Public Out As String
        Public Err As String
    End Structure

    Private Shared Function Run(stdin As String, ParamArray args As String()) As RunResult
        Dim out = New StringWriter()
        Dim err = New StringWriter()
        Dim code = Cli.Run(args, New StringReader(If(stdin, "")), out, err)
        Return New RunResult With {.Code = code, .Out = Lf(out.ToString()), .Err = Lf(err.ToString())}
    End Function

    Private Shared ReadOnly ModelPath As String = DataPath("model.xml")

    <Fact>
    Public Sub No_arguments_prints_help()
        Dim r = Run(Nothing)
        Assert.Equal(0, r.Code)
        Assert.Contains("Usage:", r.Out)
    End Sub

    <Fact>
    Public Sub Version()
        Dim r = Run(Nothing, "-ver")
        Assert.Equal(0, r.Code)
        Assert.StartsWith("objParser 5.", r.Out)
    End Sub

    <Fact>
    Public Sub Single_bare_argument_is_a_path_over_stdin()
        Dim r = Run("<a n=""1""/><a n=""2""/>", "//a/@n")
        Assert.Equal(0, r.Code)
        Assert.Equal("1" & vbLf & "2" & vbLf, r.Out)
        Assert.Equal("", r.Err) ' no map warning noise for built-in maps
    End Sub

    <Fact>
    Public Sub File_input_sorting_and_count()
        Dim r = Run(Nothing, "-f", ModelPath, "-x", "//flags/flag/@attr", "-s", "-count")
        Assert.Equal("apfile" & vbLf & "idxA" & vbLf & "other" & vbLf & "portfolio_index" & vbLf & "textnode" & vbLf, r.Out)
        Assert.Contains("Results returned: 6", r.Err)
    End Sub

    <Fact>
    Public Sub Several_paths_are_concatenated_in_order()
        Dim r = Run(Nothing, "-f", ModelPath, "-x", "//decision/@name", "-x", "//variable/@name")
        Assert.Equal("d1" & vbLf & "alpha" & vbLf & "beta" & vbLf & "gamma" & vbLf, r.Out)
    End Sub

    <Fact>
    Public Sub Delete_with_replace()
        Dim r = Run("<r><a/><b/><a></a></r>", "-x", "//a", "-delete", "-replace", "-")
        Assert.Equal("<r>-<b/>-</r>", r.Out)
    End Sub

    <Fact>
    Public Sub Delete_with_several_paths_applies_each()
        ' v4.1 printed an empty document here
        Dim r = Run("<r><a/><b/><c/></r>", "-x", "//a", "-x", "//c", "-delete")
        Assert.Equal("<r><b/></r>", r.Out)
    End Sub

    <Fact>
    Public Sub Clean_then_query()
        Dim r = Run(Nothing, "-f", ModelPath, "-clean", "//projatt", "-clean", "//index", "-x", "//@name")
        Assert.Equal("demo alpha beta gamma d1 o1", r.Out.Trim().Replace(vbLf, " "))
    End Sub

    <Fact>
    Public Sub Manual_delimiters_with_quote_token()
        Dim r = Run("x=""1"" y=""2""", "-d", "=~", "-D", "~")
        Assert.Equal("1" & vbLf & "2" & vbLf, r.Out)
    End Sub

    <Fact>
    Public Sub Built_in_map_switches()
        Assert.Equal("Invoice" & vbLf & "Customer" & vbLf, Run(Nothing, "-vb", "-f", DataPath("Sample.vb"), "-x", "//:").Out)
        Assert.Equal("Invoice" & vbLf & "InvoiceLine" & vbLf, Run(Nothing, "-cs", "-f", DataPath("Sample.cs"), "-x", "//:").Out)
    End Sub

    <Fact>
    Public Sub External_map_file()
        Dim r = Run(Nothing, "-m", DataPath("mediawiki.map"), "-f", DataPath("release.wiki"), "-x", "//h2")
        Assert.Equal(0, r.Code)
        Assert.Equal("Steps" & vbLf & "Owners" & vbLf, r.Out)
    End Sub

    <Fact>
    Public Sub Show_map_prints_map_text()
        Dim r = Run(Nothing, "-showmap", "-cs")
        Assert.Equal(0, r.Code)
        Assert.Contains("!=%{};", r.Out)
        Assert.Contains("#=#region\s", r.Out)
    End Sub

    <Fact>
    Public Sub Map_file_warnings_go_to_stderr()
        Dim mapPath = Path.Combine(Path.GetTempPath(), $"objparser-{Guid.NewGuid():N}.map")
        File.WriteAllText(mapPath, "broken line" & vbLf & "a=<" & vbLf & "a=>" & vbLf & "a=I")
        Try
            Dim r = Run("<x>", "-m", mapPath, "-x", "//a")
            Assert.Equal(0, r.Code)
            Assert.Equal("x" & vbLf, r.Out)
            Assert.Contains("warning:", r.Err)
        Finally
            File.Delete(mapPath)
        End Try
    End Sub

    <Fact>
    Public Sub Timings_on_stderr()
        Dim r = Run(Nothing, "-f", ModelPath, "-x", "//variable/@name", "-t")
        Assert.Contains("Timings", r.Err)
        Assert.Contains("variable", r.Err)
        Assert.DoesNotContain("Timings", r.Out)
    End Sub

    <Theory>
    <InlineData(1, "-f", "no-such-file.xml", "-x", "//a")>
    <InlineData(1, "-m", "no-such.map", "-x", "//a")>
    <InlineData(2, "-f", "x.xml", "-x")>
    <InlineData(2, "-bogus", "-x", "//a")>
    <InlineData(2, "-f", "x.xml")>
    <InlineData(2, "-x", "//a", "-replace", "z")>
    Public Sub Errors_have_messages_and_exit_codes(expected As Integer, ParamArray args As String())
        Dim r = Run("", args)
        Assert.Equal(expected, r.Code)
        Assert.StartsWith("objParser:", r.Err)
        Assert.Equal("", r.Out)
    End Sub

End Class
