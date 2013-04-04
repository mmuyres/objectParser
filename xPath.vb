
Imports System.Text

Namespace ObjectParser


    ' -------------------------------------------------------------------------------------------
    ' xPath Object to handle sub parsing
    ' -------------------------------------------------------------------------------------------
    ' 3 entries per token def, /open/close/result scope/  
    'prefix=oTok[field][jump+nTok]
    '  reserved chars
    '  * = replace by field name
    '  ^ = jump
    '  | = split
    '  
    '  other requests
    '  I = inner results
    '  O = outer results

    ' ------ example --------
    ' =<*^>
    ' =/>?/>|</*^>
    ' =I|O
    '!=<*
    '!=/>?/>|>
    '!=I|I
    '@=*="
    '@="
    '@=I


    Public Class xPath
        Private strPath As String
        Private pathPosition As Integer

        Private subPath As String
        Private subPathPosition As Integer

        Private pathList() As String
        Private pathPrefix() As String
        Private columnList() As String
        Public prefixList() As String

        Private m_TokenType As ObjectParser.ParserTokenType

        Private hasInit As Boolean
        Private hasSubPath As Boolean

        Private globalSearch As Boolean
        Private rootSearch As Boolean

        Private InnerFilter As Boolean
        Private lastMultiPath As String

        Private pathListTimming() As Double

        Public Sub New(ByVal tokenTypeDef As ObjectParser.ParserTokenType)
            Me.hasInit = False
            Me.hasSubPath = False
            Me.pathPosition = -1
            Me.globalSearch = False
            Me.rootSearch = False
            Me.m_TokenType = New ObjectParser.ParserTokenType
            tokenTypeDef.Clone(Me.m_TokenType)
            lastMultiPath = ""
        End Sub


        Public ReadOnly Property Initalized() As Boolean
            Get
                Return Me.hasInit
            End Get
        End Property
        Public ReadOnly Property Position() As Integer
            Get
                Return Me.pathPosition
            End Get
        End Property

        Public ReadOnly Property EOF() As Boolean
            Get
                Dim rtn As Boolean
                rtn = False
                If Me.pathList Is Nothing Then
                    rtn = True
                ElseIf Me.pathList.Length = 0 Then
                    rtn = True
                ElseIf (Me.pathPosition >= Me.pathList.Length) Then
                    rtn = True
                ElseIf Me.pathList(Me.pathPosition) Is Nothing Then
                    rtn = True
                End If
                Return rtn
            End Get
        End Property

        Public ReadOnly Property isPathGlobal() As Boolean
            Get
                Return Me.globalSearch
            End Get
        End Property
        Public ReadOnly Property isPathRoot() As Boolean
            Get
                Return Me.rootSearch
            End Get
        End Property

        ' we are assuming the prefixList was manually set or a Path has been properly Init(x,y) already
        Public Property path() As String
            Get
                Return strPath
            End Get
            Set(ByVal Value As String)
                Me.Init(Value)
            End Set
        End Property

        Public Property multipath() As String
            Get
                Return Me.lastMultiPath
            End Get
            Set(ByVal value As String)
                Me.lastMultiPath = value
            End Set
        End Property


        Sub Init(ByVal pathVal As String, ByVal prefixListArray As String())
            Me.prefixList = prefixListArray
            Me.Init(pathVal)
        End Sub

        ' must have a prefix list before initializing a path
        Sub Init(ByVal pathVal As String)
            Dim subOpen As Integer, subLength As Integer, hasFilters As Boolean
            hasFilters = False
            Me.pathPosition = -1
            Me.m_TokenType.Init()

            If pathVal.StartsWith("//") Then
                Me.globalSearch = True
                Me.strPath = pathVal.Substring(2)
            ElseIf pathVal.StartsWith("/") Then
                Me.rootSearch = True
                Me.strPath = pathVal.Substring(1)
            Else
                Me.strPath = pathVal
            End If

            If pathVal.IndexOf(Me.m_TokenType.XPATH_FILTER_OPEN) > -1 Then
                hasFilters = True
            Else
                hasFilters = False
            End If

            ' try to find a sub path
            subOpen = Me.strPath.IndexOf(Me.m_TokenType.XPATH_SUB_OPEN)
            If subOpen > -1 And Me.strPath.IndexOf(Me.m_TokenType.XPATH_SUB_CLOSE) > -1 Then
                subOpen += Me.m_TokenType.XPATH_SUB_OPEN.Length
                subLength = Me.strPath.IndexOf(Me.m_TokenType.XPATH_SUB_CLOSE) - subOpen
                Me.subPath = Me.strPath.Substring(subOpen, subLength)
                Me.strPath = Me.strPath.Replace(Me.subPath, Me.m_TokenType.XPATH_SUB_REF)
                hasSubPath = True
            Else
                Me.subPath = ""
            End If

            ' split path into an array
            If hasFilters = False Then
                If Me.strPath.IndexOf(Me.m_TokenType.XPATH_SPLIT) > -1 Then
                    Me.pathList = Me.strPath.Split(Me.m_TokenType.XPATH_SPLIT)
                    ReDim Me.pathPrefix(Me.pathList.Length)
                    ReDim Me.pathListTimming(Me.pathList.Length)
                Else
                    ReDim Me.pathList(1)
                    ReDim Me.pathPrefix(1)
                    ReDim Me.pathListTimming(1)
                    Me.pathList(0) = Me.strPath
                End If
            Else
                ' replace / with #:# for so we can split and keep the filter entry
                Dim i As Integer, idxPath As Char(), pathBuilder As StringBuilder, inFilter As Integer
                pathBuilder = New StringBuilder
                idxPath = Me.strPath.ToCharArray()
                inFilter = 0
                For i = 0 To idxPath.Length - 1
                    If Not inFilter And idxPath(i) = Me.m_TokenType.XPATH_FILTER_OPEN Then
                        '                        pathBuilder.Append(idxPath(i))
                        inFilter += 1
                    End If
                    If inFilter And idxPath(i) = Me.m_TokenType.XPATH_FILTER_CLOSE Then
                        '                       pathBuilder.Append(idxPath(i))
                        inFilter -= 1
                    End If
                    If inFilter = 0 And idxPath(i) = "/" Then
                        pathBuilder.Append(Me.m_TokenType.XPATH_TOKEN_TEMP_DELIM)
                    Else
                        pathBuilder.Append(idxPath(i))
                    End If
                Next
                Me.strPath = pathBuilder.ToString()
                Me.pathList = Me.strPath.Split(Me.m_TokenType.XPATH_TOKEN_TEMP_DELIM)
                ReDim Me.pathPrefix(Me.pathList.Length)
                ReDim Me.pathListTimming(Me.pathList.Length)
            End If

            Me.ProcessPathList()

            ' find subpath position
            Me.subPathPosition = -1
            If Me.subPath.Length > 0 Then
                Dim i As Integer
                For i = 0 To Me.pathList.Length - 1
                    If Me.pathList(i).IndexOf(Me.m_TokenType.XPATH_SUB_REF) <> -1 Then
                        Me.subPathPosition = i
                        Exit For
                    End If
                Next
            End If
            hasInit = True
        End Sub

        Public Sub ProcessPathList()
            ' process the path list
            Dim i As Integer
            For i = 0 To Me.pathList.Length - 1
                If Me.pathList(i) Is Nothing Then
                    Me.pathPrefix(i) = " "
                Else
                    ' check for prefix on path entry
                    Me.pathPrefix(i) = Me.GetPathPrefix(Me.pathList(i))
                End If
                'if there is a non-default prefix, then remove it from the path - unless it is the current result prefix
                If Me.pathPrefix(i).Length > 0 AndAlso Me.pathPrefix(i) <> " " AndAlso Not Me.pathList(i).StartsWith(Me.m_TokenType.CURRENT_RESULT_TOKEN) AndAlso Me.pathList(i) <> Me.m_TokenType.XPATH_FIELD_TOKEN Then
                    Me.pathList(i) = Me.pathList(i).Substring(Me.pathPrefix(i).Length)
                End If
            Next
        End Sub
        ' lookup the prefix for a path 
        Public Function GetPathPrefix(ByVal pathValue As String) As String
            Dim i As Integer
            If Not pathValue Is Nothing Then
                For i = 0 To Me.prefixList.Length - 1
                    '                    Console.Error.WriteLine("--Prefix: " & Me.prefixList(i))
                    If Not Me.prefixList(i) Is Nothing AndAlso pathValue.StartsWith(Me.prefixList(i)) Then
                        Return Me.prefixList(i)
                    End If
                Next
            End If
            Return " "
        End Function

        ' fetch next path
        Public Function Fetch() As Boolean
            If Me.pathPosition < Me.pathList.Length Then
                Me.pathPosition += 1
                Me.pathListTimming(Me.pathPosition) = DateTime.Now.Ticks
                If Me.pathPosition > 0 Then
                    Me.pathListTimming(Me.pathPosition - 1) = Me.getTimeSpan(Me.pathListTimming(Me.pathPosition - 1), DateTime.Now.Ticks)
                End If
                Return True
            Else
                Me.pathListTimming(Me.pathPosition) = Me.getTimeSpan(Me.pathListTimming(Me.pathPosition), DateTime.Now.Ticks)
                Return False
            End If
        End Function
        ' get current path token
        Public Function GetTokenField() As String
            If Me.pathPosition > -1 AndAlso Me.pathPosition < Me.pathList.Length AndAlso Not Me.pathList(Me.pathPosition) Is Nothing Then
                Dim n As Integer
                n = Me.pathList(Me.pathPosition).IndexOf(Me.m_TokenType.XPATH_FILTER_OPEN)
                If n > -1 AndAlso Me.pathList(Me.pathPosition) <> Me.m_TokenType.XPATH_FIELD_TOKEN Then
                    Return Me.pathList(Me.pathPosition).Substring(0, n)
                Else
                    Return Me.pathList(Me.pathPosition)
                End If
            Else
                Return ""
            End If
        End Function
        Public Function GetTokenPrefix() As String
            If Me.pathPosition > -1 And Me.pathPosition < Me.pathList.Length Then
                '                Console.Error.WriteLine("--xPath Token Prefix: " & Me.pathPrefix(pathPosition))
                Return Me.pathPrefix(pathPosition)
            Else
                Return " "
            End If
        End Function
        Public Function getMutiPathList() As String()
            Dim mList As String(), mPath As String
            mPath = Me.getMultiPath()
            If mPath.Length > 0 Then
                mList = mPath.Split(Me.m_TokenType.MULTIPATH_DELIM)
            Else
                ReDim mList(1)
                mList(0) = ""
            End If
            Return mList
        End Function
        Private Function getMultiPath() As String
            Dim mPath As String, pathSeg As String, mOpen As Integer, mClose As Integer, mLen As Integer
            mPath = ""
            pathSeg = Me.GetTokenField()
            mOpen = pathSeg.IndexOf(Me.m_TokenType.MULTIPATH_OPEN) + Me.m_TokenType.MULTIPATH_OPEN.Length
            mClose = pathSeg.IndexOf(Me.m_TokenType.MULTIPATH_CLOSE)
            mLen = mClose - mOpen
            If mLen > 0 Then
                mPath = pathSeg.Substring(mOpen, mLen)
            End If
            If mPath = Me.m_TokenType.MULTIPATH_PREV Then
                mPath = lastMultiPath
            ElseIf mPath.Length > 0 Then
                Me.lastMultiPath = mPath
            End If
            Return mPath
        End Function
        Public Function isMultiPath() As Boolean
            Dim rtn As Boolean, pathSeg As String
            rtn = False
            pathSeg = Me.GetTokenField()
            If pathSeg.IndexOf(Me.m_TokenType.MULTIPATH_OPEN) <> -1 AndAlso pathSeg.IndexOf(Me.m_TokenType.MULTIPATH_CLOSE) <> -1 Then
                rtn = True              
            End If
            Return rtn
        End Function
        Public Function hasFilter() As Boolean
            If Not Me.EOF Then
                Return (Me.pathList(Me.pathPosition).IndexOf(Me.m_TokenType.XPATH_FILTER_OPEN) > -1)
            Else
                Return False
            End If
        End Function
        Public ReadOnly Property isInnerFilter() As Boolean
            Get
                Return Me.InnerFilter
            End Get
        End Property
        Public Function GetTokenFilter() As String
            Me.InnerFilter = False
            If Me.pathPosition > -1 And Me.pathPosition < Me.pathList.Length Then
                Dim n As Integer, nn As Integer, szFCLOSE As Integer, tkFilter As String

                n = Me.pathList(Me.pathPosition).IndexOf(Me.m_TokenType.XPATH_FILTER_OPEN)
                nn = Me.pathList(Me.pathPosition).LastIndexOf(Me.m_TokenType.XPATH_FILTER_CLOSE)
                szFCLOSE = Me.m_TokenType.XPATH_FILTER_CLOSE.Length

                'lastPos = n + 1
                'While True
                '    nn = Me.pathList(Me.pathPosition).IndexOf(Me.m_TokenType.XPATH_FILTER_CLOSE, lastPos)


                '    If nn = -1 Then
                '        nn = lastPos - 1
                '        Exit While
                '    End If
                '    If Me.pathList(Me.pathPosition).IndexOf(Me.m_TokenType.XPATH_FILTER_OPEN, lastPos) = -1 Then
                '        If nn = Me.pathList(Me.pathPosition).Length OrElse Me.pathList(Me.pathPosition).IndexOf(Me.m_TokenType.XPATH_FILTER_CLOSE, nn) = -1 Then
                '            Exit While
                '        Else
                '            lastPos = Me.pathList(Me.pathPosition).IndexOf(Me.m_TokenType.XPATH_FILTER_CLOSE, lastPos) + 1
                '        End If
                '    Else
                '        lastPos = Me.pathList(Me.pathPosition).IndexOf(Me.m_TokenType.XPATH_FILTER_OPEN, lastPos) + 1
                '    End If
                'End While


                If n > -1 And nn > -1 Then
                    tkFilter = Me.pathList(Me.pathPosition).Substring(n + szFCLOSE, nn - n - szFCLOSE)
                    If tkFilter.StartsWith("-") Then
                        tkFilter = tkFilter.Substring(1)
                        Me.InnerFilter = True
                    ElseIf tkFilter.StartsWith("+") Then
                        tkFilter = tkFilter.Substring(1)
                        Me.InnerFilter = False
                    End If
                    Return tkFilter
                Else
                    Return ""
                End If
            Else
                Return ""
            End If
        End Function

        Public Function hasSubToken() As Boolean
            If Me.subPath.Length > 0 Then
                Return True
            Else
                Return False
            End If
        End Function
        Public Function GetSubToken() As String
            Return Me.subPath
        End Function

        Public Function UpdateSubToken(ByVal SubTokenResults As String) As Boolean
            If SubTokenResults Is Nothing Or Me.subPathPosition = -1 Then
                Return False
            Else
                Me.pathList(Me.subPathPosition) = Me.pathList(Me.subPathPosition).Replace(Me.m_TokenType.XPATH_SUB_OPEN & Me.m_TokenType.XPATH_SUB_REF & Me.m_TokenType.XPATH_SUB_CLOSE, SubTokenResults.Trim())
                Return True
            End If
        End Function

        Public Function getPathTimes() As String
            Dim statsList As StringBuilder, i As Integer, maxWidthPath, maxWidthTime, maxWidthTime2 As Integer
            statsList = New StringBuilder
            maxWidthPath = 10 : maxWidthTime = 4 : maxWidthTime2 = 4
            ' find max path width
            For i = 0 To Me.pathList.Length - 1
                If Not Me.pathList(i) Is Nothing Then
                    If maxWidthPath < Me.pathList(i).Length Then
                        maxWidthPath = Me.pathList(i).Length
                    End If
                End If
                If maxWidthTime < Me.pathListTimming(i).ToString().Split(".")(0).Length() Then
                    maxWidthTime = Me.pathListTimming(i).ToString().Split(".")(0).Length()
                End If
                If maxWidthTime2 < Me.pathListTimming(i).ToString().Split(".")(1).Length() Then
                    maxWidthTime2 = Me.pathListTimming(i).ToString().Split(".")(1).Length()
                End If
            Next
            maxWidthPath += 5
            If maxWidthPath < 15 Then maxWidthPath = 15
            ' create timing output
            statsList.AppendLine()
            statsList.Append(" ", (maxWidthPath - 12) / 2)
            statsList.Append("Path Segment")
            statsList.Append(" ", (maxWidthPath - 12) / 2)
            statsList.Append("     Time")
            statsList.AppendLine()
            For i = 0 To Me.pathList.Length - 1
                If Not Me.pathList(i) Is Nothing Then
                    statsList.Append(" ", maxWidthPath - Me.pathList(i).Length)
                Else
                    statsList.Append(" ", maxWidthPath)
                End If
                statsList.Append(Me.pathList(i))
                statsList.Append("  -  ")
                statsList.Append(" ", maxWidthTime - Me.pathListTimming(i).ToString().Split(".")(0).Length())
                statsList.Append(Me.pathListTimming(i))
                statsList.Append(" ", maxWidthTime2 - Me.pathListTimming(i).ToString().Split(".")(1).Length())
                statsList.Append(" ms")
                statsList.AppendLine()
            Next
            Return statsList.ToString()
        End Function

        Private Function getTimeSpan(ByVal tStart As Long, ByVal tEnd As Long) As Double
            Dim tDif As Double
            If tEnd = 0 Or tStart = 0 Then
                tDif = 0
            Else
                tDif = (New TimeSpan(tEnd).Subtract(New TimeSpan(tStart)).TotalMilliseconds)
            End If
            Return tDif
        End Function

    End Class

End Namespace