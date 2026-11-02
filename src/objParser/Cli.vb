Imports System.IO
Imports System.Reflection
Imports ObjectParser

Namespace ObjParserCli

    ''' <summary>Options gathered from the command line.</summary>
    Friend NotInheritable Class CliOptions
        Public Property Paths As New List(Of String)
        Public Property CleanPaths As New List(Of String)
        Public Property InputFile As String
        Public Property MapFile As String
        Public Property MapName As String = "xml"
        Public Property OpenDelimiter As String
        Public Property CloseDelimiter As String
        Public Property Sort As Boolean
        Public Property DeleteMode As Boolean
        Public Property KeepPath As String
        Public Property Replacement As String
        Public Property ShowTime As Boolean
        Public Property ShowCount As Boolean
        Public Property ShowMap As Boolean
        Public Property ShowVersion As Boolean
        Public Property ShowHelp As Boolean
        Public Property Verbose As Boolean
        Public Property Debug As Boolean
        Public Property ResultsToken As String
    End Class

    Public Module Cli

        Public Const ExitOk As Integer = 0
        Public Const ExitInputError As Integer = 1
        Public Const ExitUsageError As Integer = 2

        ''' <summary>Runs the tool. Returns the process exit code.</summary>
        Public Function Run(args As String(), input As TextReader, out As TextWriter, err As TextWriter) As Integer
            Dim options As CliOptions = Nothing
            Dim problem = ParseArguments(args, options)
            If problem IsNot Nothing Then
                err.WriteLine("objParser: " & problem)
                err.WriteLine("Run 'objParser -help' for usage.")
                Return ExitUsageError
            End If

            If options.ShowHelp Then
                WriteUsage(out)
                Return ExitOk
            End If
            If options.ShowVersion Then
                out.WriteLine(VersionLine())
                Return ExitOk
            End If

            Dim syntax = New ParserSyntax()
            If options.ResultsToken IsNot Nothing Then syntax.CurrentResultToken = options.ResultsToken

            Dim map As TokenMap
            Try
                If options.MapFile IsNot Nothing Then
                    If Not File.Exists(options.MapFile) Then
                        err.WriteLine($"objParser: token map not found: {options.MapFile}")
                        Return ExitInputError
                    End If
                    map = TokenMap.FromFile(options.MapFile, syntax)
                Else
                    map = TokenMap.BuiltIn(options.MapName, syntax)
                End If
            Catch ex As ArgumentException
                err.WriteLine("objParser: " & ex.Message)
                Return ExitUsageError
            End Try

            ' built-in maps carry a known, harmless warning; show warnings for user maps or on request
            If options.MapFile IsNot Nothing OrElse options.Verbose Then
                For Each w In map.Warnings
                    err.WriteLine("warning: " & w)
                Next
            End If

            If options.ShowMap Then
                out.WriteLine(If(options.MapFile IsNot Nothing, $"Token map: {options.MapFile}", $"Built-in token map: {options.MapName}"))
                out.WriteLine()
                out.Write(map.ToString())
                Return ExitOk
            End If

            Dim text As String
            If options.InputFile IsNot Nothing Then
                If Not File.Exists(options.InputFile) Then
                    err.WriteLine($"objParser: file not found: {options.InputFile}")
                    Return ExitInputError
                End If
                text = File.ReadAllText(options.InputFile)
            Else
                text = input.ReadToEnd()
            End If

            Dim started = Stopwatch.GetTimestamp()
            Dim parser = New ObjectParser.Parser(map, text) With {
                .Log = err, .Verbose = options.Verbose, .Debug = options.Debug
            }
            If options.CleanPaths.Count > 0 Then parser.Clean(options.CleanPaths)

            Dim count As Integer
            If options.DeleteMode Then
                For Each p In options.Paths
                    parser.Delete(p, options.Replacement, options.KeepPath)
                    count += parser.LastMatches.Count
                Next
                out.Write(parser.Text)
            ElseIf options.Paths.Count > 0 Then
                Dim results = parser.QueryAll(options.Paths)
                count = results.Count
                WriteResults(out, results, options.Sort)
            Else
                Dim results = parser.FindBetween(options.OpenDelimiter, options.CloseDelimiter)
                count = results.Count
                WriteResults(out, results, options.Sort)
            End If

            If options.ShowTime Then
                err.WriteLine("---------- Timings ----------")
                err.Write(parser.Statistics.ToString())
                err.WriteLine($"Total: {Stopwatch.GetElapsedTime(started).TotalMilliseconds.ToString("0.000", Globalization.CultureInfo.InvariantCulture)} ms")
            End If
            If options.ShowCount Then err.WriteLine($"Results returned: {count}")
            Return ExitOk
        End Function

        Private Sub WriteResults(out As TextWriter, results As IReadOnlyList(Of String), sort As Boolean)
            Dim lines As IEnumerable(Of String) = results
            If sort Then lines = results.Distinct(StringComparer.Ordinal).OrderBy(Function(r) r, StringComparer.Ordinal)
            For Each line In lines
                out.WriteLine(line)
            Next
        End Sub

        ''' <summary>Fills <paramref name="options"/>; returns an error message or Nothing.</summary>
        Friend Function ParseArguments(args As String(), ByRef options As CliOptions) As String
            options = New CliOptions()
            If args.Length = 0 Then
                options.ShowHelp = True
                Return Nothing
            End If
            ' a single bare argument is a path read against stdin
            If args.Length = 1 AndAlso Not args(0).StartsWith("-"c) Then
                options.Paths.Add(args(0))
                Return Nothing
            End If

            Dim i = 0
            While i < args.Length
                Dim arg = args(i)
                Dim value As String = Nothing
                Select Case arg
                    Case "-x", "-m", "-f", "-d", "-D", "-keep", "-replace", "-clean", "-results_token"
                        If i + 1 >= args.Length Then Return $"option {arg} needs a value"
                        value = args(i + 1)
                        i += 1
                End Select

                Select Case arg
                    Case "-x" : options.Paths.Add(value)
                    Case "-clean" : options.CleanPaths.Add(value)
                    Case "-f" : options.InputFile = value
                    Case "-m" : options.MapFile = value
                    Case "-d" : options.OpenDelimiter = value
                    Case "-D" : options.CloseDelimiter = value
                    Case "-keep" : options.KeepPath = value
                    Case "-replace" : options.Replacement = value
                    Case "-results_token" : options.ResultsToken = value
                    Case "-xml", "-xml_ext", "-html", "-xhtml", "-vb", "-cs", "-csharp" : options.MapName = arg.Substring(1)
                    Case "-s" : options.Sort = True
                    Case "-delete" : options.DeleteMode = True
                    Case "-t" : options.ShowTime = True
                    Case "-count" : options.ShowCount = True
                    Case "-showmap" : options.ShowMap = True
                    Case "-v" : options.Verbose = True
                    Case "-debug" : options.Debug = True
                    Case "-ver", "-version", "--version" : options.ShowVersion = True
                    Case "-h", "-help", "--help", "-?" : options.ShowHelp = True
                    Case Else
                        Return $"unknown option '{arg}'"
                End Select
                i += 1
            End While

            If options.ShowHelp OrElse options.ShowVersion OrElse options.ShowMap Then Return Nothing
            Dim hasDelimiters = options.OpenDelimiter IsNot Nothing AndAlso options.CloseDelimiter IsNot Nothing
            If options.Paths.Count = 0 AndAlso Not hasDelimiters Then
                Return "give a path with -x, or both -d and -D"
            End If
            If options.DeleteMode AndAlso options.Paths.Count = 0 Then Return "-delete needs at least one -x path"
            If (options.KeepPath IsNot Nothing OrElse options.Replacement IsNot Nothing) AndAlso Not options.DeleteMode Then
                Return "-keep and -replace only apply with -delete"
            End If
            Return Nothing
        End Function

        Private Function VersionLine() As String
            Dim asm = GetType(ObjectParser.Parser).Assembly
            Dim version = asm.GetCustomAttribute(Of AssemblyInformationalVersionAttribute)()?.InformationalVersion
            Return $"objParser {If(version, asm.GetName().Version?.ToString())}"
        End Function

        Friend Sub WriteUsage(out As TextWriter)
            out.WriteLine(VersionLine() & " - token-map based structured text parser")
            out.WriteLine("
Usage:
  objParser [options] -x <path> [-x <path> ...]
  objParser <path>                     (reads stdin)
  objParser [options] -d <open> -D <close>

Input and maps:
  -f <file>            read this file (default: stdin)
  -xml | -xml_ext | -html | -xhtml | -vb | -cs
                       built-in token map (default: xml)
  -m <mapfile>         load a token map file instead of a built-in map
  -showmap             print the token map and exit

Queries:
  -x <path>            path to evaluate; repeat to concatenate results
  -d <open> -D <close> ad-hoc delimiters instead of a path (~ means "")
  -s                   sorted, unique output
  -results_token <t>   change the current-result path token (default .)

Editing (prints the edited text):
  -delete              remove every match of the -x paths
  -replace <text>      ...replacing each match with <text>
  -keep <path>         ...replacing each match with <path> evaluated inside it
  -clean <path>        remove matches before querying (repeatable)

Diagnostics (written to stderr):
  -count  -t  -v  -debug  -ver  -help

Path examples (xml map):
  //variable/@name                      attribute values under each variable
  //variable/-node                      inner text of node elements
  //variable[@type='number']/@name      filter: = equals, ! not equal, ~ contains
  //variable[flags]  //variable[!flags] has / has not
  //{:variable,decision:}/@name         several element names at once
  //variable[@name='{{//index/@name}}'] sub-path evaluated first

Full documentation: docs/ in the project repository.")
        End Sub

    End Module

End Namespace
