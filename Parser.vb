Imports System.Text

Namespace ObjectParser

    Public Class Parser

        Private debugMode As Boolean
        Private verboseMode As Boolean
        Private m_wildCardSupportEnabled As Boolean   ' this will be disable if a token map contains a prefix that is the same as the XPATH_FIELD_TOKEN

        Private oTokenMap As ObjectParser.TokenMap
        Private oxPath As ObjectParser.xPath

        Public SortOutput As Boolean
        Public DeleteEntries As Boolean
        Public KeepPath As String
        Public ReplaceValue As String

        Private m_lastPosition As Integer
        Private m_lastResults As Hashtable
        Private m_lastResultCount As Integer

        Private m_mapFile As String
        Private m_parseBuffer As String

        Private m_TokenType As ObjectParser.ParserTokenType
        Private m_TokenCache As ObjectParser.TokenMap

        Private m_xpath As String
        Private m_xpath_list As ArrayList       'list of xpaths to return
        Private m_clean_list As ArrayList       'xpath results to be removed from buffer
        Private m_IsClean As Boolean            'is the parse buffer clean

        Private m_timing_Filter As Double
        Private m_time_Filtering As Double

        Private m_timing_Parser As Double
        Private m_time_Parseing As Double

#Region "Initialization and Setup"

        Sub New(ByVal tokenTypeDef As ObjectParser.ParserTokenType)
            Me.m_TokenType = New ObjectParser.ParserTokenType
            tokenTypeDef.Clone(Me.m_TokenType)
            SortOutput = False
            DeleteEntries = False
            m_lastPosition = 0
            debugMode = False
            m_parseBuffer = ""
            m_mapFile = ""
            KeepPath = ""
            ReplaceValue = Nothing
            oTokenMap = New ObjectParser.TokenMap(Me.m_TokenType)
            m_TokenCache = New ObjectParser.TokenMap(Me.m_TokenType)
            oxPath = New ObjectParser.xPath(tokenTypeDef)
            m_xpath_list = New ArrayList
            m_clean_list = New ArrayList
            m_IsClean = False
            m_xpath = ""
            m_lastResultCount = 0
            m_lastResults = New Hashtable
            m_time_Filtering = 0
            m_timing_Filter = 0
            m_time_Parseing = 0
            m_timing_Parser = 0
            m_wildCardSupportEnabled = True
        End Sub
        Sub New(ByVal tokenTypeDef As ObjectParser.ParserTokenType, ByVal MapFileName As String)
            Me.New(tokenTypeDef)
            m_mapFile = MapFileName
            oTokenMap.LoadMapFile(MapFileName)
            oxPath.prefixList = Me.TokenMap.GetPreFixList()
        End Sub
        Sub New(ByVal tokenTypeDef As ObjectParser.ParserTokenType, ByVal MapFileName As String, ByRef ParseTXT As String)
            Me.New(tokenTypeDef, MapFileName)
            m_parseBuffer = ParseTXT
        End Sub

        Public Sub Clear()
            ReplaceValue = Nothing
            oTokenMap = Nothing
            m_TokenCache = Nothing
            oxPath = Nothing
            m_xpath_list = Nothing
            m_clean_list = Nothing
            Me.m_lastResults = Nothing

            SortOutput = False
            DeleteEntries = False
            m_lastPosition = 0
            debugMode = False
            m_parseBuffer = ""
            m_mapFile = ""
            KeepPath = ""
            m_IsClean = False
            m_xpath = ""
            m_lastResultCount = 0
            m_time_Filtering = 0
            m_timing_Filter = 0
            m_time_Parseing = 0
            m_timing_Parser = 0

            oTokenMap = New ObjectParser.TokenMap(Me.m_TokenType)
            m_TokenCache = New ObjectParser.TokenMap(Me.m_TokenType)
            oxPath = New ObjectParser.xPath(Me.m_TokenType)
            m_xpath_list = New ArrayList
            m_clean_list = New ArrayList
            Me.m_lastResults = New Hashtable
        End Sub

#End Region


#Region "Public Properties"
        Public Property debug() As Boolean
            Get
                Return Me.debugMode
            End Get
            Set(ByVal Value As Boolean)
                Me.debugMode = Value
            End Set
        End Property
        Public Property verbose() As Boolean
            Get
                Return Me.verboseMode
            End Get
            Set(ByVal Value As Boolean)
                Me.verboseMode = Value
            End Set
        End Property

        Public Property TokenMap() As ObjectParser.TokenMap
            Get
                Return oTokenMap
            End Get
            Set(ByVal Value As ObjectParser.TokenMap)
                oTokenMap = Value
                oxPath.prefixList = Me.TokenMap.GetPreFixList()
                Me.m_wildCardSupportEnabled = oTokenMap.isWildCardSupportEnabled
            End Set
        End Property

        ' get/set a list of xpaths to concat
        Public Property xpaths() As ArrayList
            Get
                Return Me.m_xpath_list
            End Get
            Set(ByVal value As ArrayList)
                If value Is Nothing Then
                    Me.m_xpath_list.Clear()
                Else
                    Me.m_xpath_list = value
                End If
            End Set
        End Property
        ' get/set a list of xpath to remove from source
        Public Property xcleans() As ArrayList
            Get
                Return Me.m_clean_list
            End Get
            Set(ByVal value As ArrayList)
                If value Is Nothing Then
                    Me.m_clean_list.Clear()
                Else
                    Me.m_clean_list = value
                End If
            End Set
        End Property

        ' manually set path
        Public Property path() As String
            Get
                Return Me.m_xpath
            End Get
            Set(ByVal value As String)
                If value Is Nothing Then
                    Me.m_xpath = ""
                Else
                    Me.m_xpath = value
                End If
            End Set
        End Property

        Public ReadOnly Property Count() As Integer
            Get
                Return Me.m_lastResultCount
            End Get
        End Property
        Public ReadOnly Property TimeFiltering() As Double
            Get
                Return Me.m_time_Filtering
            End Get
        End Property
        Public ReadOnly Property TimeParseing() As Double
            Get
                Return Me.m_time_Parseing
            End Get
        End Property

        Public ReadOnly Property LastResults() As Hashtable
            Get
                Return Me.m_lastResults
            End Get
        End Property
#End Region

        Public Function SetBuffer(ByVal txtString As String) As Boolean
            Me.m_lastPosition = 0
            If Not txtString Is Nothing Then
                m_parseBuffer = txtString
                m_IsClean = False
                Return True
            End If
            Return False
        End Function

        ' add an xpath to the list
        Public Sub addResultPath(ByVal xpathStr As String)
            If Not xpathStr Is Nothing AndAlso xpathStr.Length > 0 Then
                If Not Me.m_xpath_list.Contains(xpathStr) Then
                    Me.m_xpath_list.Add(xpathStr)
                End If
            End If
        End Sub
        ' add a cleaning path to list
        Public Sub addCleaningPath(ByVal xpathStr As String)
            If Not xpathStr Is Nothing AndAlso xpathStr.Length > 0 Then
                If Not Me.m_clean_list.Contains(xpathStr) Then
                    Me.m_clean_list.Add(xpathStr)
                End If
            End If
        End Sub

        ' clean out results from source prior to main parsing
        Public Sub cleanBuffer()
            If Not Me.m_IsClean Then
                Dim tmpBuffer As String, curXpath As String
                Dim tmpParser As ObjectParser.Parser
                If Me.m_clean_list.Count > 0 Then
                    tmpBuffer = Me.m_parseBuffer
                    tmpParser = New ObjectParser.Parser(Me.m_TokenType)
                    tmpParser.TokenMap = Me.TokenMap
                    tmpParser.DeleteEntries = True
                    For Each curXpath In Me.m_clean_list
                        tmpParser.SetBuffer(tmpBuffer)
                        tmpBuffer = tmpParser.xPath(curXpath)
                    Next
                    Me.m_time_Filtering += tmpParser.TimeFiltering
                    Me.m_parseBuffer = tmpBuffer
                End If
                Me.m_IsClean = True
            End If
        End Sub
#Region "Top Level Parsing routines"

        Public Function xPath(ByVal strPath As String, Optional ByVal filterValue As Hashtable = Nothing, Optional ByVal filterOperand As String = Nothing) As String
            Dim tempBuffer As Hashtable, fieldName As String, prevTokenVal As String
            Dim tokenPair As ObjectParser.TokenPair, prevTokenPair As ObjectParser.TokenPair
            Dim resultBuffer As StringBuilder
            If Me.verbose = True Then
                Console.Out.WriteLine("Path:" & strPath)
            End If
            If Not Me.m_IsClean Then
                Me.cleanBuffer()
            End If
            Me.m_lastResultCount = 0
            m_xpath = strPath       ' save path
            prevTokenPair = Nothing     ' assigned after first token pair is processed (we are past position 0)
            resultBuffer = New StringBuilder
            oxPath.Init(strPath, Me.TokenMap.GetPreFixList())
            ' handle sub paths before processing xpath
            If oxPath.hasSubToken Then
                Dim tmpParser As ObjectParser.Parser, subPathVal As String
                tmpParser = New ObjectParser.Parser(Me.m_TokenType)
                tmpParser.TokenMap = Me.TokenMap
                tmpParser.SetBuffer(Me.m_parseBuffer)
                subPathVal = tmpParser.xPath(oxPath.GetSubToken())
                Me.m_time_Filtering += tmpParser.TimeFiltering
                oxPath.UpdateSubToken(subPathVal)
            End If

            tempBuffer = New Hashtable
            fieldName = "" : prevTokenVal = "" : Me.m_lastPosition = -1
            While (oxPath.Fetch() And Not oxPath.EOF)
                fieldName = oxPath.GetTokenField()
                If Me.debugMode Then
                    Console.Out.WriteLine("fieldName: " & fieldName)
                End If

                If Me.TokenMap.PrefixExists(oxPath.GetTokenPrefix()) Then
                    tokenPair = Me.TokenMap.GetTokenPair(oxPath.GetTokenPrefix())
                    tokenPair.debug = Me.debugMode
                    If oxPath.Position = 0 Then
                        If oxPath.isMultiPath() Then
                            tempBuffer = Me.findMultiSubText(Me.m_parseBuffer, oxPath.getMutiPathList(), oxPath.GetTokenPrefix())
                        Else
                            tempBuffer = Me.FindSubText(Me.m_parseBuffer, fieldName, tokenPair)
                        End If
                    Else
                        If prevTokenPair.IO = "O" Then       ' for outer scope results jump past opening tag
                            Me.m_lastPosition = prevTokenPair.openToken.firstTagVal.Length
                        Else
                            Me.m_lastPosition = 0
                        End If
                        If oxPath.isMultiPath() Then
                            tempBuffer = Me.findMultiSubText(tempBuffer, oxPath.getMutiPathList(), oxPath.GetTokenPrefix())
                        Else
                            tempBuffer = Me.FindSubText(tempBuffer, fieldName, tokenPair)
                        End If
                    End If
                    ' track/cache the object after use
                    prevTokenPair = tokenPair
                End If

            End While
            If Not filterValue Is Nothing Then
                tempBuffer = Me.filterOutput(tempBuffer, filterValue, filterOperand)
            End If
            Me.m_lastResultCount = tempBuffer.Count
            Me.m_lastResults = tempBuffer
            Return Me.generateOutput(tempBuffer)
        End Function

        ' run multiple xpaths before returning results
        Public Function xPath(ByVal xpaths As ArrayList) As String
            Dim resultBuffer As Hashtable, pathCNT As Integer, curXpath As String, resultCNT As Integer
            resultBuffer = New Hashtable
            pathCNT = 0 : resultCNT = 0
            For Each curXpath In Me.m_xpath_list
                pathCNT += 1
                resultBuffer.Add(pathCNT, Me.xPath(curXpath))
                resultCNT += Me.m_lastResultCount
            Next
            Me.m_lastResultCount = resultCNT
            Me.m_lastResults = resultBuffer
            Return Me.generateOutput(resultBuffer)
        End Function

        ' return based on internals
        Public Function xPath() As String
            If Me.m_xpath_list.Count > 1 Then
                Return Me.xPath(Me.m_xpath_list)
            ElseIf Me.m_xpath_list.Count = 1 And Me.m_xpath.Length = 0 Then
                Return Me.xPath(Me.m_xpath_list)
            Else
                Return Me.xPath(Me.m_xpath)
            End If
        End Function

        ' process multi-path path segment in a hash buffer.  
        Public Function findMultiSubText(ByVal txtBuffer As Hashtable, ByVal fieldNames As String(), ByVal tokenPrefix As String) As Hashtable
            Dim resultsBuffer As Hashtable, tmpResults As Hashtable, resultCNT As Integer, curFieldName As String, curLastPost As Integer, curTxtBuffer As String
            Dim hashWalkerT As System.Collections.IDictionaryEnumerator, hashWalkerB As System.Collections.IDictionaryEnumerator, i As Integer, ii As Integer
            Dim tokenPair As ObjectParser.TokenPair
            resultCNT = 0 : curLastPost = Me.m_lastPosition
            resultsBuffer = New Hashtable
            ' fetch results for each path segment in list for each buffer entry
            hashWalkerB = txtBuffer.GetEnumerator()
            For ii = 1 To txtBuffer.Count
                hashWalkerB.MoveNext()
                curTxtBuffer = hashWalkerB.Value
                For Each curFieldName In fieldNames
                    If curFieldName.Length > 0 Then
                        tokenPair = Me.TokenMap.GetTokenPair(tokenPrefix)
                        Me.m_lastPosition = curLastPost
                        tmpResults = Me.FindSubText(curTxtBuffer, curFieldName, tokenPair)
                        hashWalkerT = tmpResults.GetEnumerator()
                        ' save up each result into one master list
                        For i = 1 To tmpResults.Count
                            hashWalkerT.MoveNext()
                            resultsBuffer.Add(resultCNT, hashWalkerT.Value)
                            resultCNT += 1
                        Next
                    End If
                Next
            Next
            Return resultsBuffer
        End Function
        ' process multi-path path segment in a txt buffer.  
        Public Function findMultiSubText(ByVal txtBuffer As String, ByVal fieldNames As String(), ByVal tokenPrefix As String) As Hashtable
            Dim resultsBuffer As Hashtable, tmpResults As Hashtable, resultCNT As Integer, curFieldName As String, curLastPost As Integer
            Dim hashWalker As System.Collections.IDictionaryEnumerator, i As Integer, tokenPair As ObjectParser.TokenPair
            resultCNT = 0 : curLastPost = Me.m_lastPosition
            resultsBuffer = New Hashtable
            ' fetch results for each path segment in list
            For Each curFieldName In fieldNames
                Me.m_lastPosition = curLastPost
                If curFieldName.Length > 0 Then
                    tokenPair = Me.TokenMap.GetTokenPair(tokenPrefix)
                    tmpResults = Me.FindSubText(txtBuffer, curFieldName, tokenPair)
                    hashWalker = tmpResults.GetEnumerator()
                    ' save up each result into one master list
                    For i = 1 To tmpResults.Count
                        hashWalker.MoveNext()
                        resultsBuffer.Add(resultCNT, hashWalker.Value)
                        resultCNT += 1
                    Next
                End If
            Next
            Return resultsBuffer
        End Function



#End Region

#Region "Low Level Parsing routines"

        ' command line raw token implementation
        Public Function FindSubText(ByVal openToken As String, ByVal closeToken As String) As String
            Dim ocPair As TokenPair
            Dim resultBuffer As Hashtable
            Dim objTokenType As New ObjectParser.ParserTokenType
            ocPair = New TokenPair(objTokenType, openToken.Replace("~", """"), closeToken.Replace("~", """"))
            resultBuffer = Me.FindSubText(Me.m_parseBuffer, "", ocPair)
            Me.m_lastResultCount = resultBuffer.Count
            Me.m_lastResults = resultBuffer
            Return Me.generateOutput(resultBuffer)
        End Function

        ' token pair object parsing implementation
        Public Function FindSubText(ByRef txtBuffer As Hashtable, ByVal FieldName As String, ByRef tokenPair As ObjectParser.TokenPair) As Hashtable
            Dim n As Integer, i As Integer, curLastPosition As Integer
            Dim newBuffer As Hashtable, tmpBuffer As Hashtable, hashWalker As System.Collections.IDictionaryEnumerator, tmpNewBuffer As String
            Dim tmpHashWalker As System.Collections.IDictionaryEnumerator, tmpTokenPair As ObjectParser.TokenPair

            newBuffer = New Hashtable
            tmpNewBuffer = ""
            tmpTokenPair = Nothing
            ' keep track of last position so we can reset it each text block
            curLastPosition = Me.m_lastPosition
            hashWalker = txtBuffer.GetEnumerator()
            For i = 1 To txtBuffer.Count
                Me.m_lastPosition = curLastPosition
                hashWalker.MoveNext()
                tmpTokenPair = tokenPair
                tmpBuffer = Me.FindSubText(hashWalker.Value, FieldName, tmpTokenPair.openToken, tmpTokenPair.closeToken)
                tmpHashWalker = tmpBuffer.GetEnumerator()

                For n = 1 To tmpBuffer.Count
                    tmpHashWalker.MoveNext()
                    newBuffer.Add((i * 10000) + n, tmpHashWalker.Value)
                Next
            Next
            Return newBuffer
        End Function


        ' token pair object parsing implementation
        Public Function FindSubText(ByRef txtBuffer As String, ByVal FieldName As String, ByRef tokenPair As ObjectParser.TokenPair) As Hashtable
            ' force a copy of the object, then save it for global use and prev state ref
            'Me.curTokenPair = tokenPair
            Return Me.FindSubText(txtBuffer, FieldName, tokenPair.openToken, tokenPair.closeToken)
        End Function

        ' generic parser
        Public Function FindSubText(ByVal txtBuffer As String, ByVal FieldName As String, ByRef openToken As ObjectParser.ParserToken, ByRef closeToken As ObjectParser.ParserToken) As Hashtable
            Dim outputBuffer As Hashtable, parseLength As Integer, tempBuffer As String

            Dim tParser As ObjectParser.Parser
            Dim tmpFilter As String, tmpFilterLT As String, tmpFilterRT As String, tmpFilterLTVAL As String, tmpFilterRTVAL As Hashtable
            Dim filterModeOn As Boolean, filterMatch As Boolean, filterOperand As String = ""
            Dim fltrOpenToken, fltrCloseToken As ObjectParser.ParserToken
            filterModeOn = False : filterMatch = False : tParser = Nothing
            tmpFilterLT = "" : tmpFilterRTVAL = New Hashtable : tmpFilterRT = "" : fltrOpenToken = Nothing : fltrCloseToken = Nothing

            outputBuffer = New Hashtable
            openToken.FieldValue = FieldName
            closeToken.FieldValue = FieldName
            tempBuffer = ""
            Me.m_timing_Parser = Date.Now.Ticks()
            Me.m_lastPosition -= 1
            Do
                ' check token to see if it has participated in wildcard field name replacement
                If openToken.Parent.isWildCarding Then
                    openToken.FieldValue = ""
                    closeToken.FieldValue = ""
                    Me.m_lastPosition -= 1
                End If
                ' reset tokens per iteration
                'openToken.Reset()
                closeToken.Reset()

                '                Me.m_lastPosition -= 1
                ' get first token location
                If Me.ProcessToken(openToken, txtBuffer) Then
                    ' have the closeToken eval the results from the openToken
                    closeToken.EvalCondition(openToken)
                    If Me.ProcessToken(closeToken, txtBuffer) Then
                        If Me.debugMode Then
                            Console.Out.WriteLine("Close Token isOuterResult: " & closeToken.Type.isOuterResults())
                            Console.Out.WriteLine(" Token Parent: " & closeToken.Parent.IO())
                        End If
                        ' return either the outer or inner scope
                        If closeToken.Type.isOuterResults Then
                            ' get inner text results
                            outputBuffer.Add(Me.m_lastPosition, Me.GetOuterText(txtBuffer, openToken, closeToken))
                            parseLength = closeToken.OuterPos - openToken.OuterPos
                            If Me.debugMode Then
                                Console.Out.WriteLine("OuterResults parseLength: " & parseLength)
                                If parseLength < 128 Then
                                    Console.Out.WriteLine("    txt:" & outputBuffer.Item(Me.m_lastPosition))
                                End If
                            End If
                        Else
                            outputBuffer.Add(Me.m_lastPosition, Me.GetInnerText(txtBuffer, openToken, closeToken))
                            parseLength = closeToken.InnerPos - openToken.InnerPos
                            If Me.debugMode Then
                                Console.Out.WriteLine("InnerResults parseLength: " & parseLength)
                                If parseLength < 128 Then
                                    Console.Out.WriteLine("    txt:" & outputBuffer.Item(Me.m_lastPosition))
                                End If
                            End If
                        End If
                        If parseLength <= 0 Then
                            Exit Do
                        End If
                    Else
                        Exit Do
                    End If
                Else
                    If openToken.Parent.IsCurrentResult Then
                        If Me.m_lastPosition <= 0 Then
                            Me.m_lastPosition = 1
                        End If
                        outputBuffer.Add(Me.m_lastPosition, txtBuffer)
                    End If
                    Exit Do
                End If
            Loop
            Me.m_time_Parseing += Me.getTimeSpan(Me.m_timing_Parser)

            ' filter results if filter exists
            If oxPath.hasFilter() Then
                Me.m_timing_Filter = Date.Now.Ticks()

                filterModeOn = True
                tParser = New ObjectParser.Parser(Me.m_TokenType)
                tParser.TokenMap = Me.TokenMap
                tParser.SortOutput = Me.SortOutput
                tParser.oxPath.multipath = oxPath.multipath
                tmpFilter = oxPath.GetTokenFilter()
                ' split filter into RT and LT
                If tmpFilter.IndexOf("=") > 1 Then
                    filterOperand = "="
                    tmpFilterLT = tmpFilter.Split(filterOperand)(0)
                    tmpFilterRT = tmpFilter.Split(filterOperand)(1)
                ElseIf tmpFilter.IndexOf("!") > 1 Then
                    filterOperand = "!"
                    tmpFilterLT = tmpFilter.Split(filterOperand)(0)
                    tmpFilterRT = tmpFilter.Split(filterOperand)(1)
                    If tmpFilterRT.Length = 0 Then
                        filterOperand = "N"
                    End If
                ElseIf tmpFilter.IndexOf("~") > 1 Then
                    filterOperand = "~"
                    tmpFilterLT = tmpFilter.Split(filterOperand)(0)
                    tmpFilterRT = tmpFilter.Split(filterOperand)(1)
                Else
                    ' check for existence
                    tmpFilterLT = tmpFilter
                    filterOperand = "X"
                End If
                If tmpFilterRT.StartsWith("'") And tmpFilterRT.EndsWith("'") Then
                    tmpFilterRTVAL.Add(0, tmpFilterRT.Substring(1, tmpFilterRT.Length - 2))
                Else
                    If tmpFilterRT.StartsWith("//") Then
                        tParser.SetBuffer(Me.m_parseBuffer)
                    Else
                        tParser.SetBuffer(" " & txtBuffer & " ")
                    End If
                    tParser.xPath(tmpFilterRT)
                    tmpFilterRTVAL = tParser.LastResults
                End If
                If Me.debugMode Or Me.verboseMode Then
                    If tmpFilterRTVAL.Count > 1 Then
                        Console.Out.WriteLine(" Filter Match: " & tmpFilterLT & " = ResultList of [" & tmpFilterRTVAL.Count & "]")
                    Else
                        Console.Out.WriteLine(" Filter Match: " & tmpFilterLT & " = " & tmpFilterRTVAL.Item(0).ToString())
                    End If
                End If

                ' walk results list to see if we have a match
                Dim hashWalker As System.Collections.IDictionaryEnumerator, i As Integer, newBuffer As New Hashtable
                hashWalker = outputBuffer.GetEnumerator()
                For i = 1 To outputBuffer.Count
                    hashWalker.MoveNext()
                    ' check filter condition on text block
                    If filterModeOn Then
                        tParser.SetBuffer(" " & hashWalker.Value & " ")
                        tmpFilterLTVAL = tParser.xPath(tmpFilterLT)
                        If filterOperand = "=" Then
                            If tmpFilterRTVAL.Count = 1 AndAlso (tmpFilterLTVAL.Replace(vbCrLf, "").ToLower() = tmpFilterRTVAL.Item(0).ToString().Replace(vbCrLf, "").ToLower()) Then
                                filterMatch = True
                            ElseIf tmpFilterLTVAL.Length > 1 Then
                                'secondary lookup, incase we are looking inside a block of text as a result
                                tmpFilterLTVAL = tParser.xPath(tmpFilterLT, tmpFilterRTVAL, filterOperand)
                                If tmpFilterLTVAL.Length > 0 Then
                                    filterMatch = True
                                Else
                                    filterMatch = False
                                End If
                            Else
                                filterMatch = False
                            End If
                        ElseIf filterOperand = "!" Then
                            If tmpFilterRTVAL.Count = 1 AndAlso (tmpFilterLTVAL.Replace(vbCrLf, "") <> tmpFilterRTVAL.Item(0).ToString().Replace(vbCrLf, "")) Then
                                filterMatch = True
                            ElseIf tmpFilterRTVAL.Count > 1 Then
                                'list lookup, incase we are looking inside a block of text as a result
                                tmpFilterLTVAL = tParser.xPath(tmpFilterLT, tmpFilterRTVAL, filterOperand)
                                If tmpFilterLTVAL.Length = 0 Then
                                    filterMatch = True
                                Else
                                    filterMatch = False
                                End If
                            Else
                                filterMatch = False
                            End If
                        ElseIf filterOperand = "~" Then         ' contains check
                            If tmpFilterRTVAL.Count = 1 AndAlso (tmpFilterLTVAL.Replace(vbCrLf, "").ToLower().IndexOf(tmpFilterRTVAL.Item(0).ToString().Replace(vbCrLf, "").ToLower()) <> -1) Then
                                filterMatch = True
                            ElseIf tmpFilterRTVAL.Count > 1 Then
                                'list lookup, incase we are looking inside a block of text as a result
                                tmpFilterLTVAL = tParser.xPath(tmpFilterLT, tmpFilterRTVAL, filterOperand)
                                If tmpFilterLTVAL.Length = 0 Then
                                    filterMatch = True
                                Else
                                    filterMatch = False
                                End If
                            Else
                                filterMatch = False
                            End If
                        ElseIf filterOperand = "X" Then
                            If tmpFilterLTVAL.Length > 0 Then
                                filterMatch = True
                            Else
                                filterMatch = False
                            End If
                        ElseIf filterOperand = "N" Then     ' not exists or is empty
                            If tmpFilterLTVAL.Length = 0 Then
                                filterMatch = True
                            Else
                                filterMatch = False
                            End If
                        Else
                            filterMatch = False
                        End If
                    End If
                    If filterMatch Then
                        newBuffer.Add((10000 + i), hashWalker.Value)
                    End If
                Next
                outputBuffer = newBuffer
                Me.m_time_Filtering += tParser.TimeFiltering
                Me.m_time_Filtering += Me.getTimeSpan(Me.m_timing_Filter) - tParser.TimeFiltering
                Me.m_time_Parseing += tParser.TimeParseing
            End If
            Return outputBuffer
        End Function

        Private Function GetInnerText(ByRef txtBuffer As String, ByRef openToken As ObjectParser.ParserToken, ByRef closeToken As ObjectParser.ParserToken) As String
            Dim sz As Integer
            sz = closeToken.InnerPos - openToken.InnerPos
            '    closeToken.Parent.IO = "I"
            If sz > 0 And txtBuffer.Length > openToken.OuterPos + sz Then
                Return txtBuffer.Substring(openToken.InnerPos, sz)
            End If
            Return ""
        End Function

        Private Function GetOuterText(ByRef txtBuffer As String, ByRef openToken As ObjectParser.ParserToken, ByRef closeToken As ObjectParser.ParserToken) As String
            Dim sz As Integer
            sz = closeToken.OuterPos - openToken.OuterPos
            '      closeToken.Parent.IO = "O"
            If sz > 0 And txtBuffer.Length >= openToken.OuterPos + sz Then
                Return txtBuffer.Substring(openToken.OuterPos, sz)
            End If
            Return ""
        End Function
        Private Function ProcessToken(ByRef token As ObjectParser.ParserToken, ByRef txtBuffer As String) As Boolean
            Dim rtn As Boolean
            rtn = False
            ' process only if eval has not forced a skip
            If token.Parent.IsCurrentResult Then
                rtn = False
                Me.m_lastPosition = -1
            ElseIf token.Skip Then
                rtn = True              'there is nothing to do we are already there, prev token placed us at our start point
            Else
                'try to find first token+field
                If Me.m_lastPosition <= 0 Then
                    Me.m_lastPosition = 0
                End If

                ' if we have a field based token and the field is empty, try to derive the field name
                If Me.m_wildCardSupportEnabled AndAlso token.Type.isOpenToken Then
                    If token.FieldIsWild Then
                        Dim fldStart, fldEnd, curLoc As Integer, tFLD As String, fLen As Integer, startsWithToken As Boolean, applyOffset As Integer
                        Do
                            startsWithToken = False
                            applyOffset = 0
                            ' if the buffer starts with the first tag then start at the begining, else move in the tag length to avoid double matching
                            If txtBuffer.StartsWith(token.firstTagVal) Then
                                curLoc = 0
                                startsWithToken = True
                            Else
                                curLoc = Me.m_lastPosition + token.firstTagVal.Length - 1
                            End If
                            ' if we start on non-blank, find blank or token first 
                            If curLoc < txtBuffer.Length AndAlso Not Me.isEmpty(txtBuffer.Substring(curLoc, 1)) Then
                                curLoc = Me.GetNextWhiteSpace(txtBuffer, token.firstTagVal, curLoc)
                            End If
                            If startsWithToken Then
                                fldStart = 0
                            Else
                                fldStart = Me.GetNextNonWhiteSpace(txtBuffer, curLoc)
                            End If
                            ' check to see if the we are sitting on the tokens first char
                            If fldStart + token.firstTagVal.Length <= txtBuffer.Length AndAlso txtBuffer.Substring(fldStart, token.firstTagVal.Length).StartsWith(token.firstTagVal) Then
                                fldStart += 1
                                applyOffset = 1
                            End If
                            If token.nextTagToken.Length = 0 Then
                                fldEnd = Me.GetNextWhiteSpace(txtBuffer, token.firstTagVal, fldStart)
                            Else
                                fldEnd = Me.GetNextWhiteSpace(txtBuffer, token.nextTagVal, fldStart)
                            End If
                            If fldStart < 0 Then
                                fldStart = 0
                            End If
                            fLen = fldEnd - fldStart
                            If fLen > txtBuffer.Length - fldStart Then
                                fLen = txtBuffer.Length - fldStart
                            End If
                            If fldEnd > fldStart Then
                                tFLD = txtBuffer.Substring(fldStart, fLen)
                                token.Parent.openToken.FieldValue = tFLD
                                token.Parent.closeToken.FieldValue = tFLD
                                ' sanity check the field value we have extrapolated
                                If fldStart + token.firstTagVal.Length <= txtBuffer.Length AndAlso txtBuffer.Substring(fldStart - applyOffset, token.firstTagVal.Length) = token.firstTagVal Then
                                    token.Parent.isWildCarding = True
                                    ' back up if we have adjusted the field
                                    If Me.m_lastPosition > token.Parent.openToken.FieldValue.Length Then
                                        Me.m_lastPosition -= token.Parent.openToken.FieldValue.Length
                                    End If
                                    Exit Do
                                Else
                                    token.Parent.openToken.FieldValue = ""
                                    token.Parent.closeToken.FieldValue = ""
                                    ' advance to the next possible starting point
                                    If Me.isEmpty(txtBuffer.Substring(Me.m_lastPosition, 1)) Then
                                        Me.m_lastPosition = GetNextNonWhiteSpace(txtBuffer, token.firstTagVal, Me.m_lastPosition)
                                    Else
                                        Me.m_lastPosition = GetNextWhiteSpace(txtBuffer, token.firstTagVal, Me.m_lastPosition)
                                    End If
                                End If
                            End If
                            If Me.m_lastPosition >= txtBuffer.Length Then
                                Exit Do
                            End If
                            Me.m_lastPosition += 1      ' always move forward to avoid infinite loop
                        Loop
                    End If
                End If
                ' ensure we do not nav left to far
                If Me.m_lastPosition < 0 Then
                    Me.m_lastPosition = 0
                End If

                '' if we have a close token that is conditional search and the first token are the same, skip.
                'If token.Type.isCloseToken() AndAlso token.isSearchConditional AndAlso token.Parent.openToken.firstTagToken = token.Parent.closeToken.firstTagToken Then
                '    Me.m_lastPosition = 0
                'End If

                ' find first tag
                If txtBuffer.Length > 0 Then
                    token.firstTagPos = txtBuffer.IndexOf(token.firstTagVal, Me.m_lastPosition)
                Else
                    token.firstTagPos = 0
                End If
                ' check and continue parsing secondary part of tag
                If Not token.hadError And token.firstPosEnd <= txtBuffer.Length Then
                    Me.m_lastPosition = token.firstPosEnd
                    If token.firstPosEnd + token.NextPrevCharsSize > txtBuffer.Length Then
                        token.NextPrevCharsSize = txtBuffer.Length - token.firstPosEnd
                    End If
                    token.firstTagNextChars = txtBuffer.Substring(token.firstPosEnd, token.NextPrevCharsSize)
                    token.NextTagPrevChars = ""
                    If Me.debugMode Then
                        Console.Out.WriteLine("First Tag Val: " & token.firstTagVal & "  Tag Pos: " & token.firstTagPos & " Fist Tag Prev Chars:" & token.firstTagNextChars)
                    End If
                    rtn = True

                    ' search a list of tokens to find - forward or backwards
                    If token.isSearchConditional Then
                        Dim maxPOS, minPOS, srchPOS As Integer, srchVAL As String, minTOKVAL, maxTOKVAL As String
                        If txtBuffer.Length = 0 Then
                            token.nextTagPos = 0
                        Else
                            minPOS = txtBuffer.Length : srchPOS = 0 : maxPOS = 0 : minTOKVAL = "" : maxTOKVAL = ""
                            For Each srchVAL In token.nextTagConditions
                                srchPOS = txtBuffer.IndexOf(srchVAL, Me.m_lastPosition)
                                If minPOS > srchPOS Then
                                    minPOS = srchPOS
                                    minTOKVAL = srchVAL
                                End If
                                If maxPOS < srchPOS Then
                                    maxPOS = srchPOS
                                    maxTOKVAL = srchVAL
                                End If
                            Next
                            If token.Type.seekBackward Then
                                token.nextTagCondVal = minTOKVAL
                                token.nextTagPos = minPOS
                            ElseIf token.Type.seekForward Then
                                token.nextTagCondVal = minTOKVAL
                                token.nextTagPos = maxPOS
                            End If
                            token.PrevNextTokenSize = token.nextTagCondVal.Length
                        End If
                        If Not token.hadError Then
                            Me.m_lastPosition = token.nextPosEnd
                            token.NextTagPrevChars = txtBuffer.Substring(token.nextPosEnd - token.PrevNextTokenSize, token.PrevNextTokenSize)
                            If token.nextPosEnd + token.NextPrevCharsSize > txtBuffer.Length Then
                                token.NextPrevCharsSize = txtBuffer.Length - token.firstPosEnd
                            End If
                            If Me.debugMode Then
                                Console.Out.WriteLine("Next Tag Val: " & token.nextTagCondVal & "  Tag Pos: " & token.nextTagPos & " Next Tag Prev Chars:" & token.NextTagPrevChars)
                            End If
                        End If

                    End If

                    ' off-set when last char of first token is same as a begining of the next token open
                    If token.Type.isOpenToken Then
                        If token.nextTagVal.Length = 0 Then
                            If token.firstTagVal.EndsWith(token.Parent.closeToken.firstTagVal) Then
                                Me.m_lastPosition += 1
                            End If
                        Else
                            If token.firstTagVal.EndsWith(token.nextTagVal) Then
                                Me.m_lastPosition += 1
                            End If
                        End If
                    End If

                    ' check for jump
                    If token.hasJump Then
                        If token.firstTagVal.EndsWith(token.nextTagVal.Substring(0, 1)) Then    ' off-set when last char of first token is same as a single char next token
                            Me.m_lastPosition += 1
                        End If
                        token.nextTagPos = txtBuffer.IndexOf(token.nextTagVal, Me.m_lastPosition - 1)
                        If Not token.hadError Then
                            Me.m_lastPosition = token.nextPosEnd
                            token.NextTagPrevChars = txtBuffer.Substring(token.nextPosEnd - token.NextPrevCharsSize, token.NextPrevCharsSize)
                            If Me.debugMode Then
                                Console.Out.WriteLine("Next Tag Val: " & token.nextTagVal & "  Tag Pos: " & token.nextTagPos & " Next Tag Prev Chars:" & token.NextTagPrevChars)
                            End If
                        End If
                    End If
                End If
                If Me.debugMode = True OrElse (Me.verbose AndAlso token.hadError()) Then
                    Console.Error.WriteLine("Token Had Error:" & token.nextTagVal)
                End If
            End If
            Return rtn
        End Function

#End Region

#Region "Low level navigation"

        ' find the next white space from current location
        Protected Function GetNextWhiteSpace(ByVal curTxtBuffer As String, ByVal altValCheck As String, ByVal startPos As Integer) As Integer
            Dim curLoc, maxLoc As Integer, curChar As String, altValCheckClean As String
            curLoc = startPos
            maxLoc = curTxtBuffer.Length
            altValCheckClean = altValCheck.Replace(Me.m_TokenType.XPATH_FIELD_TOKEN, "")
            While curLoc < maxLoc
                curChar = curTxtBuffer.Substring(curLoc, 1)
                If Me.isEmpty(curChar) OrElse altValCheckClean.StartsWith(curChar) Then
                    '' if we match all the way up to the token part given, then we should return where we startin searching
                    'If altValCheckClean.StartsWith(curChar) Then
                    '    curLoc = startPos + altValCheckClean.Length
                    'End If
                    Exit While
                ElseIf curLoc < curTxtBuffer.Length Then
                    curLoc += 1
                End If
            End While
            Return curLoc
        End Function
        Protected Function GetNextWhiteSpace(ByVal curTxtBuffer As String, ByVal startPos As Integer) As Integer
            Dim curLoc, maxLoc As Integer, curChar As String
            curLoc = startPos
            maxLoc = curTxtBuffer.Length
            While curLoc < maxLoc
                curChar = curTxtBuffer.Substring(curLoc, 1)
                If Me.isEmpty(curChar) Then
                    Exit While
                ElseIf curLoc < curTxtBuffer.Length Then
                    curLoc += 1
                End If
            End While
            Return curLoc
        End Function

        ' find the next non white space from current location, allow check for token value also.  Find next non-while or next token value
        Protected Function GetNextNonWhiteSpace(ByVal curTxtBuffer As String, ByVal altValCheck As String, ByVal startPos As Integer) As Integer
            Dim curLoc, maxLoc As Integer, curChar As String
            curLoc = startPos
            maxLoc = curTxtBuffer.Length
            While curLoc < maxLoc
                curChar = curTxtBuffer.Substring(curLoc, 1)
                If (Me.isEmpty(curChar) OrElse altValCheck.StartsWith(curChar)) AndAlso curLoc < curTxtBuffer.Length Then
                    curLoc += 1
                Else
                    Exit While
                End If
            End While
            Return curLoc
        End Function
        Protected Function GetNextNonWhiteSpace(ByVal curTxtBuffer As String, ByVal startPos As Integer) As Integer
            Dim curLoc, maxLoc As Integer, curChar As String
            curLoc = startPos
            maxLoc = curTxtBuffer.Length
            While curLoc < maxLoc
                curChar = curTxtBuffer.Substring(curLoc, 1)
                If Me.isEmpty(curChar) Then
                    curLoc += 1
                Else
                    Exit While
                End If
            End While
            ' move past cur pos if we have not moved.
            If curLoc = startPos AndAlso curLoc < curTxtBuffer.Length Then
                curLoc += 1
            End If
            Return curLoc
        End Function

#End Region

#Region "Output Control Routines"

        ' take a hashtable of results and make the output string
        Private Function generateOutput(ByVal parseResults As Hashtable) As String
            Dim hashWalker As System.Collections.IDictionaryEnumerator, resultBuffer As StringBuilder, i As Integer
            resultBuffer = New StringBuilder
            If DeleteEntries Then
                ' sort the results to reduce the number of replacements
                Dim tParser As New Parser(Me.m_TokenType)
                tParser.TokenMap = Me.TokenMap
                tParser.SortOutput = Me.SortOutput
                Dim sortArray As String()
                ReDim sortArray(parseResults.Count)
                hashWalker = parseResults.GetEnumerator()
                For i = 0 To parseResults.Count - 1
                    hashWalker.MoveNext()
                    sortArray(i) = hashWalker.Value
                Next
                Array.Sort(sortArray)
                For i = 0 To sortArray.Length - 2
                    If (Not sortArray(i) Is Nothing) And (sortArray(i) <> sortArray(i + 1)) Then
                        ' remove entry from initial data glob or replace with sub select
                        If KeepPath.Length > 0 Then
                            tParser.SetBuffer(sortArray(i))
                            m_parseBuffer = m_parseBuffer.Replace(sortArray(i), tParser.xPath(Me.KeepPath))
                        ElseIf Not ReplaceValue Is Nothing Then
                            m_parseBuffer = m_parseBuffer.Replace(sortArray(i), ReplaceValue)
                        Else
                            m_parseBuffer = m_parseBuffer.Replace(sortArray(i), "")
                        End If
                    End If
                Next
                If Not sortArray(sortArray.Length - 1) Is Nothing Then
                    If KeepPath.Length > 0 Then
                        tParser.SetBuffer(sortArray(sortArray.Length - 1))
                        m_parseBuffer = m_parseBuffer.Replace(sortArray(sortArray.Length - 1), tParser.xPath(Me.KeepPath))
                    ElseIf Not ReplaceValue Is Nothing Then
                        m_parseBuffer = m_parseBuffer.Replace(sortArray(sortArray.Length - 1), ReplaceValue)
                    Else
                        m_parseBuffer = m_parseBuffer.Replace(sortArray(sortArray.Length - 1), "")
                    End If
                End If
                Me.m_time_Filtering += tParser.TimeFiltering
                Me.m_time_Parseing += tParser.TimeParseing
                resultBuffer.Append(m_parseBuffer)
            ElseIf SortOutput Then
                Dim sortArray As String()
                ReDim sortArray(parseResults.Count)
                hashWalker = parseResults.GetEnumerator()
                For i = 0 To parseResults.Count - 1
                    hashWalker.MoveNext()
                    sortArray(i) = hashWalker.Value
                Next
                Array.Sort(sortArray)
                For i = 0 To sortArray.Length - 2
                    If (Not sortArray(i) Is Nothing) And (sortArray(i) <> sortArray(i + 1)) Then
                        resultBuffer.Append(sortArray(i))
                        resultBuffer.Append(vbCrLf)
                    End If
                Next
                If Not sortArray(sortArray.Length - 1) Is Nothing Then
                    resultBuffer.Append(sortArray(sortArray.Length - 1))
                End If
            Else
                Dim txtA As String, txtB As String
                hashWalker = parseResults.GetEnumerator()
                txtA = ""
                txtB = ""
                For i = 1 To parseResults.Count
                    If i > 1 Then
                        txtA = hashWalker.Value
                    End If
                    hashWalker.MoveNext()
                    txtB = hashWalker.Value
                    If txtA <> txtB Then
                        resultBuffer.Append(hashWalker.Value)
                        resultBuffer.Append(vbCrLf)
                    End If
                Next
            End If
            Return resultBuffer.ToString()
        End Function


        ' filter a hashtable of results (case insensitive searching)
        Private Function filterOutput(ByVal parseResults As Hashtable, ByVal filterValue As Hashtable, ByVal filterOperand As String) As Hashtable
            Dim hashWalkerP As System.Collections.IDictionaryEnumerator, resultBuffer As Hashtable, i As Integer
            Dim hashWalkerF As System.Collections.IDictionaryEnumerator, ii As Integer
            Dim hashValP As String, hashValF As String, filterMatched As Boolean

            resultBuffer = New Hashtable
            hashWalkerP = parseResults.GetEnumerator()
            hashWalkerF = filterValue.GetEnumerator()
            Me.m_timing_Filter = Date.Now.Ticks()
            For i = 1 To parseResults.Count
                hashWalkerP.MoveNext()
                hashValP = hashWalkerP.Value.ToString().ToLower()

                ' walk RT hand list of values
                hashWalkerF.Reset()
                filterMatched = False
                For ii = 1 To filterValue.Count
                    hashWalkerF.MoveNext()
                    hashValF = hashWalkerF.Value.ToString.ToLower()
                    Select Case filterOperand
                        Case "="        ' match LT to any RT
                            If hashValP = hashValF Then
                                filterMatched = True
                                Exit For
                            End If
                        Case "!"        ' match LT not any RT
                            If hashValP <> hashValF Then
                                filterMatched = True
                            Else
                                filterMatched = False
                                Exit For
                            End If
                        Case "~"        ' match LT contains any RT
                            If hashValP.IndexOf(hashValF) <> -1 Then
                                filterMatched = True
                                Exit For
                            End If
                    End Select
                Next

                ' add result only once
                If filterMatched Then
                    resultBuffer.Add(hashWalkerP.Key, hashWalkerP.Value)
                End If
            Next
            Me.m_time_Filtering += Me.getTimeSpan(Me.m_timing_Filter)

            Return resultBuffer
        End Function

#End Region


#Region "Misc Support Routines"

        Private Function delWhiteSpace(ByVal str As String) As String
            Dim nStr As String
            nStr = Replace(str, Chr(9), "")
            nStr = Replace(nStr, Chr(10), "")
            nStr = Replace(nStr, Chr(13), "")
            If nStr Is Nothing Then
                nStr = ""
            End If
            Return nStr
        End Function

        Public Function getStatistics() As String
            Dim statList As StringBuilder, maxWidthTime As Integer
            statList = New StringBuilder
            maxWidthTime = 4

            If maxWidthTime < Me.m_time_Parseing.ToString.Split(".")(0).Length Then
                maxWidthTime = Me.m_time_Parseing.ToString.Split(".")(0).Length
            End If
            If maxWidthTime < Me.m_time_Filtering.ToString.Split(".")(0).Length Then
                maxWidthTime = Me.m_time_Filtering.ToString.Split(".")(0).Length
            End If

            statList.AppendLine(oxPath.getPathTimes())

            If Me.m_time_Parseing > 0 Then
                statList.Append("Time Parsing  : ")
                statList.Append(" ", maxWidthTime - Me.m_time_Parseing.ToString.Split(".")(0).Length())
                statList.AppendLine(Me.m_time_Parseing & " ms")
            End If
            If Me.m_time_Filtering > 0 Then
                statList.Append("Time Filtering: ")
                statList.Append(" ", maxWidthTime - Me.m_time_Filtering.ToString.Split(".")(0).Length())
                statList.AppendLine(Me.m_time_Filtering & " ms")
            End If

            Return statList.ToString()
        End Function

        Private Function getTimeSpan(ByVal tStart As Long, Optional ByVal tEnd As Long = 0) As Double
            Dim tDif As Double
            If tEnd = 0 Then
                tEnd = DateTime.Now.Ticks
            End If
            If tEnd = 0 Or tStart = 0 Then
                tDif = 0
            Else
                tDif = (New TimeSpan(tEnd).Subtract(New TimeSpan(tStart)).TotalMilliseconds)
            End If
            Return tDif
        End Function

        Private Function isEmpty(ByVal txtStr As String) As Boolean
            Dim rtn As Boolean = False
            Select Case txtStr
                Case " "
                    rtn = True
                Case vbTab
                    rtn = True
                Case vbLf
                    rtn = True
                Case vbCr
                    rtn = True
            End Select
            Return rtn
        End Function

#End Region

    End Class

End Namespace