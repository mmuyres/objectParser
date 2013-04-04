
Imports System.IO
Imports System.Text
Imports Microsoft.VisualBasic.Strings
Imports System.IO.FileSystemInfo
Imports System.IO.FileInfo
Imports System.IO.StreamWriter


Module Module1

    Public Sub main(ByVal CmdArgs() As String)
        Dim sTime As Long, eTime As Long, ArgNum As Integer
        Dim showTIME, showCOUNT, showVER As Boolean
        Dim tokMap As ObjectParser.TokenMap
        Dim oParser As ObjectParser.Parser

        showTIME = False : showCOUNT = False : showVER = False
        ' this definition object will be shared to allow configuration changes to propogate
        Dim tokenTypeDef As New ObjectParser.ParserTokenType
        tokenTypeDef.Init()

        Dim InternalMap As Boolean, mapType As String, SortOutput As Boolean
        InternalMap = True
        SortOutput = False
        mapType = "xml"

        Dim xPathVal As String = "", keepPathVal As String = "", replaceVal As String = Nothing, sDelim As String = "", eDelim As String = "", TokenMapFile As String = ""

        Dim fileName As String = ""
        Dim debugMode As Boolean, verboseMode As Boolean, deleteMode As Boolean, showMap As Boolean, m_reprocessArgs As Boolean
        Dim fileTXT As String, fStat As Boolean
        fStat = False : debugMode = False : verboseMode = False : deleteMode = False : showMap = False : m_reprocessArgs = False
        fileTXT = ""
        '     Console.Out.WriteLine("oparse:")

        If CmdArgs.Length = 1 Then   ' If only one arg then assume xpath with input Console.In, unless it is showmap
            If CmdArgs(0) = "-showmap" Then
                showMap = True
            ElseIf CmdArgs(0) = "-ver" Then
                showVER = True
            Else
                xPathVal = CmdArgs(0)
            End If
        End If

        If CmdArgs.Length > 1 Then            ' if there are several arguments, then process them
            For ArgNum = 0 To UBound(CmdArgs)
                '                MsgBox("-n:[" & ArgNum & "] -v:[" & CmdArgs(ArgNum) & "]")
                ' main args
                If CmdArgs(ArgNum) = "-x" Then
                    xPathVal = CmdArgs(ArgNum + 1)
                End If
                If CmdArgs(ArgNum) = "-m" Then
                    TokenMapFile = CmdArgs(ArgNum + 1)
                    InternalMap = False
                End If
                If CmdArgs(ArgNum) = "-f" Then
                    fileName = CmdArgs(ArgNum + 1)
                End If

                ' manually define a token pair for single tokenpair processing
                If CmdArgs(ArgNum) = "-d" Then
                    sDelim = CmdArgs(ArgNum + 1)
                End If
                If CmdArgs(ArgNum) = "-D" Then
                    eDelim = CmdArgs(ArgNum + 1)
                End If

                ' built-in maps
                If CmdArgs(ArgNum) = "-xml" Then
                    InternalMap = True
                    mapType = "xml"
                End If
                If CmdArgs(ArgNum) = "-xml_ext" Then
                    InternalMap = True
                    mapType = "xml_ext"
                End If
                If CmdArgs(ArgNum) = "-html" Then
                    InternalMap = True
                    mapType = "html"
                End If
                If CmdArgs(ArgNum) = "-xhtml" Then
                    InternalMap = True
                    mapType = "xml"
                End If
                If CmdArgs(ArgNum) = "-vb" Then
                    InternalMap = True
                    mapType = "vb"
                End If

                ' output controls
                If CmdArgs(ArgNum) = "-s" Then
                    SortOutput = True
                End If

                ' stripping and inverse processing
                If CmdArgs(ArgNum) = "-delete" Then
                    deleteMode = True
                End If
                If CmdArgs(ArgNum) = "-keep" Then
                    keepPathVal = CmdArgs(ArgNum + 1)
                End If
                If CmdArgs(ArgNum) = "-replace" Then
                    replaceVal = CmdArgs(ArgNum + 1)
                End If

                ' show info about parser
                If CmdArgs(ArgNum).StartsWith("-ver") Then
                    showVER = True
                End If
                If CmdArgs(ArgNum) = "-t" Then
                    showTIME = True
                End If
                If CmdArgs(ArgNum) = "-count" Then
                    showCOUNT = True
                End If
                If CmdArgs(ArgNum) = "-showmap" Then
                    showMap = True
                End If

                ' debug and verbosity
                If CmdArgs(ArgNum) = "-v" Then
                    verboseMode = True
                End If
                If CmdArgs(ArgNum) = "-debug" Then
                    debugMode = True
                End If

                ' special overloads to internal mappings
                If CmdArgs(ArgNum) = "-results_token" Then
                    tokenTypeDef.CURRENT_RESULT_TOKEN = CmdArgs(ArgNum + 1)
                End If

                ' if we need to load up more values after the parser object has been made
                If CmdArgs(ArgNum) = "-clean" OrElse (xPathVal.Length > 0 AndAlso CmdArgs(ArgNum) = "-x") Then
                    m_reprocessArgs = True
                End If

            Next ArgNum
        End If

        If CmdArgs.Length = 0 Then
            ShowUsage(True, True, True, True, True)
            Exit Sub
        ElseIf showVER Then
            ShowUsage(True, False, False, False, False)
            Exit Sub
        End If
        If showTIME Then sTime = Date.Now.Ticks

        'begin creating token map
        tokMap = New ObjectParser.TokenMap(tokenTypeDef)

        If Not InternalMap And TokenMapFile.Length > 0 Then
            tokMap.LoadMapFile(TokenMapFile)
        End If
        If InternalMap Then
            tokMap.LoadMap(makeTokenMap(mapType))
        End If

        If showMap Then
            If InternalMap Then
                Console.Out.WriteLine("Internal Token Map: [" & mapType & "]")
                Console.Out.WriteLine(" ")
                Console.Out.WriteLine(makeTokenMap(mapType))
            Else
                Console.Out.WriteLine("Loaded Token Map: [" & TokenMapFile & "]")
                Console.Out.WriteLine(" ")
                Console.Out.WriteLine(tokMap.getTokenMap())
            End If
            ShowUsage(False, False, True, False, False)
            Exit Sub
        End If

        '        If xPathVal.Length > 0 Then
        '       oxPath = New ObjectParser.xPath
        '      oxPath.Init(xPathVal, tokMap.GetPreFixList())
        '
        '       End If

        If fileName.Length = 0 Then   ' If no filename then assume Console.In
            fStat = GetFileContents(Console.OpenStandardInput(102400), fileTXT)
            'fStat = GetFileContents(Console.In, fileTXT)
        ElseIf fileName.Length >= 3 Then
            fStat = GetFileContents(fileName, fileTXT)
        Else
            fStat = False
        End If


        If fStat Then
            oParser = New ObjectParser.Parser(tokenTypeDef)
            oParser.debug = debugMode
            oParser.verbose = verboseMode
            oParser.TokenMap = tokMap
            oParser.SortOutput = SortOutput
            oParser.DeleteEntries = deleteMode
            oParser.KeepPath = keepPathVal
            oParser.ReplaceValue = replaceVal
            oParser.SetBuffer(fileTXT)

            ' process args again looking for multi-entry
            If m_reprocessArgs Then
                ' process args that can exists multiple times
                If CmdArgs.Length > 1 Then            ' if there are several arguments, then process them
                    For ArgNum = 0 To UBound(CmdArgs)
                        '                MsgBox("-n:[" & ArgNum & "] -v:[" & CmdArgs(ArgNum) & "]")
                        If CmdArgs(ArgNum) = "-clean" Then
                            oParser.addCleaningPath(CmdArgs(ArgNum + 1))
                        End If
                        If CmdArgs(ArgNum) = "-x" Then
                            oParser.addResultPath(CmdArgs(ArgNum + 1))
                        End If
                    Next ArgNum
                End If
            Else
                oParser.path = xPathVal
            End If

            '(TokenMapFile, fileTXT)
            If verboseMode = True Then
                Console.Out.WriteLine("Mode: Verbose - ON")
            End If
            If debugMode = True Then
                Console.Out.WriteLine("Mode: Debug - ON")
            End If

            If xPathVal.Length > 1 Then
                Console.Out.WriteLine(oParser.xPath())
            ElseIf sDelim.Length > 0 And eDelim.Length > 0 Then
                Console.Out.WriteLine(oParser.FindSubText(sDelim, eDelim))
            Else
                ShowUsage(True, False, False, True, True)
                Exit Sub
            End If

            If showTIME Then
                Console.Error.WriteLine("---------- Timings ----------")
                Console.Error.WriteLine(oParser.getStatistics())
            End If

            If showCOUNT Then
                Console.Error.WriteLine(" ")
                Console.Error.WriteLine("Results returned: " & oParser.Count)
            End If
        Else
            Console.Error.WriteLine("File or Stream Error or not specified")
        End If

        If showTIME Then
            eTime = DateTime.Now.Ticks
            Console.Error.WriteLine(" ")
            Console.Error.WriteLine("Total Time:" & totalTime(sTime, eTime) & " ms")
        End If
    End Sub

    Private Function GetFileContents(ByVal fName As String, ByRef fTXT As String) As Boolean
        ' Create an instance of StreamReader to read from a file.
        Dim sr As StreamReader = New StreamReader(fName)
        If File.Exists(fName) Then
            fTXT = sr.ReadToEnd
            sr.Close()
        Else
            Console.Error.WriteLine("File Not Found:" & fName)
            Return False
        End If
        Return True
    End Function

    Private Function GetFileContents(ByRef fObj As System.IO.Stream, ByRef fTXT As String) As Boolean
        ' Create an instance of StreamReader to read from a file.
        Dim tObj As TextReader = New System.IO.StreamReader(fObj, System.Text.Encoding.Default, True, 102400)
        Try
            fTXT = tObj.ReadToEnd
            Return True
        Catch ex As Exception
            fTXT = ""
            Return False
        End Try
    End Function

    Private Function GetFileContents(ByRef fObj As TextReader, ByRef fTXT As String) As Boolean
        ' Create an instance of StreamReader to read from a file.
        If Not fObj Is Nothing Then
            Try
                fTXT = fObj.ReadToEnd
            Catch es As EndOfStreamException
                fObj.Close()
                Exit Try
            Catch ex As Exception
                fObj.Close()
                Throw ex
            End Try
        End If
        Return True
    End Function

    Private Function totalTime(ByVal tStart As Long, ByVal tEnd As Long) As Double
        Dim tDif As Double
        If tEnd = 0 Or tStart = 0 Then
            tDif = 0
        Else
            tDif = (New TimeSpan(tEnd).Subtract(New TimeSpan(tStart)).TotalMilliseconds)
        End If
        Return tDif
    End Function


    Public Function makeTokenMap(ByVal MapName As String) As String
        Dim xMap As StringBuilder
        xMap = New StringBuilder

        Select Case MapName.ToLower()
            Case "xml"
                xMap.Append(" =<*^>" & vbCrLf)
                xMap.Append(" =/>?/>|</*^>" & vbCrLf)
                xMap.Append(" =O|O" & vbCrLf)

                xMap.Append("@=*=""" & vbCrLf)
                xMap.Append("@=""" & vbCrLf)
                xMap.Append("@=I" & vbCrLf)

                xMap.Append("*=</" & vbCrLf)
                xMap.Append("*=>" & vbCrLf)
                xMap.Append("*=I" & vbCrLf)

                xMap.Append("!=>" & vbCrLf)
                xMap.Append("!=<" & vbCrLf)
                xMap.Append("!=I" & vbCrLf)

                xMap.Append("$= " & vbCrLf)
                xMap.Append("$==""^""" & vbCrLf)
                xMap.Append("$=I" & vbCrLf)

                xMap.Append("-=<*^>" & vbCrLf)
                xMap.Append("-=/>?/>|</*^>" & vbCrLf)
                xMap.Append("-=I|I" & vbCrLf)


            Case "xml_ext"
                xMap.Append(" =<*^>" & vbCrLf)
                xMap.Append(" =/>?/>|</*^>" & vbCrLf)
                xMap.Append(" =O|O" & vbCrLf)

                xMap.Append("@=*=""" & vbCrLf)
                xMap.Append("@=""" & vbCrLf)
                xMap.Append("@=I" & vbCrLf)

                xMap.Append("$= " & vbCrLf)
                xMap.Append("$==""^""" & vbCrLf)
                xMap.Append("$=I" & vbCrLf)

                xMap.Append("-=<*^>" & vbCrLf)
                xMap.Append("-=/>?/>|</*^>" & vbCrLf)
                xMap.Append("-=I|I" & vbCrLf)

                xMap.Append("inner:=>" & vbCrLf)
                xMap.Append("inner:=<" & vbCrLf)
                xMap.Append("inner:=I" & vbCrLf)

                xMap.Append("node:=<*" & vbCrLf)
                xMap.Append("node:=/>?/>|>" & vbCrLf)
                xMap.Append("node:=O" & vbCrLf)

                xMap.Append("attname:= " & vbCrLf)
                xMap.Append("attname:==""^""" & vbCrLf)
                xMap.Append("attname:=I" & vbCrLf)

                xMap.Append("comment:=<!--" & vbCrLf)
                xMap.Append("comment:=-->" & vbCrLf)
                xMap.Append("comment:=O" & vbCrLf)

                xMap.Append("nodename:=<" & vbCrLf)
                xMap.Append("nodename:= | |/|>|/>" & vbCrLf)
                xMap.Append("nodename:=I|I|I" & vbCrLf)

                ' node name finder v2
                'xMap.Append("*=</" & vbCrLf)
                'xMap.Append("*=>" & vbCrLf)
                'xMap.Append("*=I" & vbCrLf)

                ' old inner method
                xMap.Append("!=>" & vbCrLf)
                xMap.Append("!=<" & vbCrLf)
                xMap.Append("!=I" & vbCrLf)


            Case "vb"
                xMap.Append(" =Class *" & vbCrLf)
                xMap.Append(" =End Class" & vbCrLf)
                xMap.Append(" =O" & vbCrLf)
                xMap.Append("!=Function *^(" & vbCrLf)
                xMap.Append("!=End Function" & vbCrLf)
                xMap.Append("!=O" & vbCrLf)
                xMap.Append("@=Sub *^(" & vbCrLf)
                xMap.Append("@=End Sub" & vbCrLf)
                xMap.Append("@=I" & vbCrLf)

                xMap.Append("#= " & vbCrLf)
                xMap.Append("#= As *" & vbCrLf)
                xMap.Append("#=I" & vbCrLf)


                xMap.Append("$=Sub " & vbCrLf)
                xMap.Append("$=(" & vbCrLf)
                xMap.Append("$=I" & vbCrLf)

                xMap.Append(":=Class " & vbCrLf)
                xMap.Append(":= " & vbCrLf)
                xMap.Append(":=I" & vbCrLf)

                xMap.Append(";=ByVal " & vbCrLf)
                xMap.Append(";= As" & vbCrLf)
                xMap.Append(";=I" & vbCrLf)

                xMap.Append("f=Function ^(" & vbCrLf)
                xMap.Append("f=)" & vbCrLf)
                xMap.Append("f=O" & vbCrLf)
                xMap.Append("s=Sub ^(" & vbCrLf)
                xMap.Append("s=)" & vbCrLf)
                xMap.Append("s=O" & vbCrLf)

                xMap.Append("p=* Function^(" & vbCrLf)
                xMap.Append("p=)" & vbCrLf)
                xMap.Append("p=O" & vbCrLf)
                xMap.Append("v=* " & vbCrLf)
                xMap.Append("v= As" & vbCrLf)
                xMap.Append("v=I" & vbCrLf)

                xMap.Append("m=*." & vbCrLf)
                xMap.Append("m=(" & vbCrLf)
                xMap.Append("m=I" & vbCrLf)

            Case Else
                xMap.Append(" =<*^>" & vbCrLf)
                xMap.Append(" =/>?/>|</*^>" & vbCrLf)
                xMap.Append(" =I|O" & vbCrLf)
                xMap.Append("@=*=""" & vbCrLf)
                xMap.Append("@=""" & vbCrLf)
                xMap.Append("@=I" & vbCrLf)
        End Select

        Return xMap.ToString()

    End Function

    Public Sub ShowUsage(ByVal showHeader As Boolean, ByVal showUsage As Boolean, ByVal showMapUsage As Boolean, ByVal showPathUsage As Boolean, ByVal showExamples As Boolean)


        If showHeader Then
            Console.Out.WriteLine(My.Application.Info.Title.ToString() & " - " & My.Application.Info.Version.ToString())
            Console.Out.WriteLine(" Author:" & My.Application.Info.CompanyName.ToString() & " - " & My.Application.Info.Copyright.ToString())
        End If

        If showUsage Then
            Console.Out.WriteLine("------------------------")
            Console.Out.WriteLine("Usage:  objParser [options] -x <xpath>")
            Console.Out.WriteLine(" ")
            Console.Out.WriteLine(" objParser -x <xpath>        -reads from pipe, if no -f option is used.")
            Console.Out.WriteLine(" -f <fileName>               -specify file to read, if omited then user stdin.")
            Console.Out.WriteLine(" -x <sudo xpath>             -specify sudo xpath to parse, can appear multiple times.")
            Console.Out.WriteLine(" -[xml|xml_ext|html|xhtml|vb]        -specify built-in token map to use, if omited then xml is used")
            Console.Out.WriteLine(" -m                          -specify token map to use, if specified built-in maps are disabled")
            Console.Out.WriteLine(" -s                          -sorted unique output")

            Console.Out.WriteLine(" ")
            Console.Out.WriteLine(" -delete                     -delete results from input")
            Console.Out.WriteLine(" -keep                       -used with delete, replace with sub path")
            Console.Out.WriteLine(" -replace                    -used with delete, replace results with string")
            Console.Out.WriteLine(" -clean                      -used to remove results from source.")

            Console.Out.WriteLine(" ")
            Console.Out.WriteLine(" -t                          -show time for parse")
            Console.Out.WriteLine(" -count                      -show count for parse")
            Console.Out.WriteLine(" -debug                      -turn debug mode on")

            Console.Out.WriteLine(" -showmap                    -display the current Token Map")
            Console.Out.WriteLine(" -results_token              -change the prefix used to return the current results")

            Console.Out.WriteLine(" ")
            Console.Out.WriteLine("--- Optional Manual Delimiters when not using xpath   ---")
            Console.Out.WriteLine("  -d <delimiter>              -specify Start Delimiter")
            Console.Out.WriteLine("  -D <delimiter>              -specify End Delimiter")
        End If

        If showMapUsage Then
            Console.Out.WriteLine(" ")
            Console.Out.WriteLine("Token Map and Delimiter Convention")
            Console.Out.WriteLine("<string> as literal")
            Console.Out.WriteLine("~ converted to double-quote")
            Console.Out.WriteLine("^ signifies; find string on left, then jump to string on right")
            Console.Out.WriteLine("* converted to xpath current path segment, not including any filters")
            Console.Out.WriteLine("   - if segment is empty * will be replaced with wildcard name, detrived from token parts")
            Console.Out.WriteLine("| used in condtional tokens to seperate the token to be used base on matching results")
            Console.Out.WriteLine("   - if used in open token, position of closest value will be used.")
            Console.Out.WriteLine("   - example:   <| |>|/> , used to find closest position to first token.  Find < then find closest of 'space' or '>' or '/>' ")
            Console.Out.WriteLine("   - if used in close token, used to compare to prev results to determine value to be used.")
            Console.Out.WriteLine("? declares a conditional token. If prev token search ends in token on left, use first OR'd token, else second token.")
            Console.Out.WriteLine("     -example:  />?/>|</*^>")
            Console.Out.WriteLine("     If prev token's inner or outer position is '/>', then use '/>' for the second token, else use '</*^>'")
        End If

        If showPathUsage Then
            Console.Out.WriteLine(" ")
            Console.Out.WriteLine("----------- xpath conventions ----------- ")
            Console.Out.WriteLine(" //   refers to top of largest block.   /   refers first block.")
            Console.Out.WriteLine(" //<prefix><string>/<prefix><string>      simple <path>")
            Console.Out.WriteLine(" //<prefix><string>[<path>(=~!)<path>]    simple path")
            Console.Out.WriteLine(" //<prefix><string>[<path>(=~!)<path>]/./<prefix><string>[<path>(=~!)<path>]    complex path with multiple filters")

            Console.Out.WriteLine(" ")
            Console.Out.WriteLine("[] declares a filter for xpath")
            Console.Out.WriteLine("     -scope:  LT side limited to parent, RT side global space.  Single quotes denote literal comparisons.")
            Console.Out.WriteLine("     -example:  //flags/flag[@attr='textnode']")
            Console.Out.WriteLine("     -example:  //variable[flags/flag/@value=//variable/@name]    (get all variables that are referenced by another variable.")
            Console.Out.WriteLine("[-] declares a inner filter for xpath, the filter check will operate on the inner results")
            Console.Out.WriteLine("     -example:  //flags[-flag/@attr='textnode']")
            Console.Out.WriteLine("[+] declares a outer filter for xpath.  Supported incase a prefix is the minus sign.")
            Console.Out.WriteLine("     -example:  //variable[+flag/@attr='textnode']")

            Console.Out.WriteLine(" ")
            Console.Out.WriteLine("     -matching logic:")
            Console.Out.WriteLine("        LT Val :  RT Val    -case insensitive single comparison")
            Console.Out.WriteLine("        LT List :  RT Val   -case insensitive comparison for each LT to RT")
            Console.Out.WriteLine("           = matches when any from LT = RT")
            Console.Out.WriteLine("           ! matches when none from LT = RT")
            Console.Out.WriteLine("           ~ matches when any from LT contains RT")
            Console.Out.WriteLine("        LT List :  RT List   -case insensitive comparison for each LT to each RT")
            Console.Out.WriteLine("           = matches when LT = any RT")
            Console.Out.WriteLine("           ! matches when LT not any RT")
            Console.Out.WriteLine("           ~ matches when LT contains any RT")
            Console.Out.WriteLine("        LT Val :  RT List  -  acts the same as above...")

            Console.Out.WriteLine(" ")
            Console.Out.WriteLine("Sub Paths - a special xpath that will be rendered prior to main xpath being parsed, only one may exist.  Wrapped in {{...}}  e.g.: {{//path[filter]/path}}  ")
            Console.Out.WriteLine("     -example:  //variable[@name='{{//index[flags/flag/@attr='portfolio_index']/@name}}']")

            Console.Out.WriteLine(" ")
            Console.Out.WriteLine("Multi Paths - special xpath segment that is a list of values wrapped in {:...:}  e.g.: //{:variable,decision,objective:}  ")
            Console.Out.WriteLine("              to reference the last multi-path use {:*:}  e.g.: //{:variable,decision,objective:}/node:{:*:}  ")
            Console.Out.WriteLine("     -example:  //{:variable,decision,objective:}[flags/flag/@='{{//index[flags/flag/@attr~'portfolio_index']/@name}}']/./node:{:*:}/@name")
            Console.Out.WriteLine("            shows any variable, decision or objective that is using the index that is flagged with 'portfolio_index'.")
            Console.Out.WriteLine(" ")

        End If

        If showExamples Then
            '//{:variable,decision,objective:}[flags/flag/@='{{//index[flags/flag/@attr~'portfolio_index']/@name}}']/./node:{:*:}
            Console.Out.WriteLine(" ")
            Console.Out.WriteLine("---      example sudo [xml] xpath for recursive parsing        ---")
            Console.Out.WriteLine("root/@name        = get the named attribute under the root node")
            Console.Out.WriteLine("root/@            = get any attributes value the root node")
            Console.Out.WriteLine("root/$            = get the names of the attributes")
            Console.Out.WriteLine("root/name         = get the name elements under the root node, returns the entire element")
            Console.Out.WriteLine("root/-name        = get the inner value of named node, returns the entire inner element")
            Console.Out.WriteLine("root/+name        = get the outer value of named node, returns the entire outer element")
            Console.Out.WriteLine("root/!            = get all attribute names")
            Console.Out.WriteLine("//projatt/att/option[@id='7']/.[@name='CNS-989']     = two filters bridge by the CURRENT_RESULTS token.")
            Console.Out.WriteLine("//objective/flags[-flag/@attr='apfile']              = two level filters using inner an filter; only filters on the parents inner element.")
            Console.Out.WriteLine("//variable[flags/flag]                               = variable that have flags.")
            Console.Out.WriteLine("//variable[!flags]                                   = variable that have no flags.")
            Console.Out.WriteLine("//choices/choice[@type=//controls/type/@value]       = left and right xpath filter.  If RT side has no //, then scope is same as LT.")

            Console.Out.WriteLine("//projatt/att[option!]/.[@snapshot_output!]          = all proj att with no options and not a snapshot_output.")
            Console.Out.WriteLine("//projatt/att[option!]/.[@snapshot_output]           = all proj att with no options, but have snapshot_output.")
            Console.Out.WriteLine("//projatt/att[option]/.[@type='pulldown']/./att[@snapshot_output~'ind']")
            Console.Out.WriteLine("                                                     = all proj att with options, of type pulldown, with a name that contains 'ind'.")

            Console.Out.WriteLine(" ")
            Console.Out.WriteLine("valid operands for filtering:  = (equals), ! (not equal),  ~ (contains)")
            Console.Out.WriteLine("                     note:  name!   =   (name is absent or empt),  nameA!nameB (nameA does not equal nameB)")
            Console.Out.WriteLine("                     note:  name    =   (name exists)")
        End If

    End Sub
End Module



