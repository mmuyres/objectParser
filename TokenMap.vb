Imports System.IO
Imports System.Text
Imports Microsoft.VisualBasic.Strings
Imports System.IO.FileSystemInfo
Imports System.IO.FileInfo
Imports System.IO.StreamWriter

Namespace ObjectParser


    ' -------------------------------------------------------------------------------------------
    ' TokenPair, holds as pair of ParserToken Objects and config info/objects
    ' -------------------------------------------------------------------------------------------

    Public Class TokenPair
        Public openToken As ObjectParser.ParserToken
        Public closeToken As ObjectParser.ParserToken
        Public tokenIO As String
        Public debug As Boolean

        Private currentIO As String
        Public m_prefixName As String
        Private toksMade As Boolean
        Private tokIsResult As Boolean  ' is this tokenPair really the current result

        Private m_TokenType As ObjectParser.ParserTokenType

        Private m_isWildCarding As Boolean

        Public Property IO() As String
            Get
                Return Me.currentIO
            End Get
            Set(ByVal Value As String)
                Me.currentIO = Value
            End Set
        End Property
        Public ReadOnly Property IsCurrentResult() As Boolean
            Get
                Return tokIsResult
            End Get
        End Property
        Public Property prefixName() As String
            Get
                Return m_prefixName
            End Get
            Set(ByVal value As String)
                m_prefixName = value
                If m_prefixName = Me.m_TokenType.CURRENT_RESULT_TOKEN Then
                    Me.tokIsResult = True
                Else
                    Me.tokIsResult = False
                End If
            End Set
        End Property
        Public Property isWildCarding() As Boolean
            Get
                Return Me.m_isWildCarding
            End Get
            Set(ByVal value As Boolean)
                Me.m_isWildCarding = value
            End Set
        End Property

        Public Sub New(ByVal tokenTypeDef As ObjectParser.ParserTokenType)
            Me.m_TokenType = New ObjectParser.ParserTokenType
            tokenTypeDef.Clone(Me.m_TokenType)
            Me.m_TokenType.Init()
            Me.Init()
        End Sub
        Public Sub New(ByVal tokenTypeDef As ObjectParser.ParserTokenType, ByVal oToken As String, ByVal cToken As String)
            Me.New(tokenTypeDef)
            Me.MakeTokens(oToken, cToken)
        End Sub

        Private Sub Init()
            tokenIO = "I"
            m_prefixName = " "
            toksMade = False
            debug = False
            tokIsResult = False
            m_isWildCarding = False
        End Sub

        Public Sub MakeTokens(ByVal oToken As String, ByVal cToken As String)
            openToken = New ObjectParser.ParserToken(Me.m_TokenType, oToken)
            openToken.SetOpenToken()
            openToken.RegisterParent(Me)
            closeToken = New ObjectParser.ParserToken(Me.m_TokenType, cToken)
            closeToken.SetCloseToken()
            closeToken.RegisterParent(Me)
            ' fixup Token IO if it is not in the same format as the tokens conditional
            If closeToken.hasCondition And tokenIO.Length < 3 Then
                tokenIO &= "|" & tokenIO
            End If
            If Not closeToken.hasCondition And tokenIO.Length > 1 Then
                tokenIO = tokenIO.Substring(0, 1)
            End If
            toksMade = True
        End Sub
        Public Sub Reset()
            If toksMade Then
                openToken.Reset()
                closeToken.Reset()
            End If
        End Sub

        Public Sub ClearBeginPos()
            If toksMade Then
                openToken.ClearBeginPos()
                closeToken.ClearBeginPos()
            End If
        End Sub
    End Class


    ' -------------------------------------------------------------------------------------------
    ' TokenMap, holds ParserToken Objects to handle defined parsing, reads token maps
    ' -------------------------------------------------------------------------------------------

    Public Class TokenMap
        Private tokenList(3) As String
        Private prefixList(3) As String

        Private tokenObjectPair As ObjectParser.TokenPair
        Private tokenObjects As Hashtable
        Private mapHasErrors As Boolean
        Private m_TokenType As ObjectParser.ParserTokenType
        Private m_wildCardSupportEnabled As Boolean

        Sub New(ByVal tokenTypeDef As ObjectParser.ParserTokenType)
            tokenObjects = New Hashtable
            mapHasErrors = False
            m_wildCardSupportEnabled = True
            Me.m_TokenType = New ObjectParser.ParserTokenType
            tokenTypeDef.Clone(Me.m_TokenType)
            Me.m_TokenType.Init()
        End Sub

        Public ReadOnly Property CurrentResultsPrefix() As String
            Get
                Return Me.m_TokenType.CURRENT_RESULT_TOKEN
            End Get
        End Property

        Public ReadOnly Property hasErrors() As Boolean
            Get
                Return mapHasErrors
            End Get
        End Property
        Public ReadOnly Property isWildCardSupportEnabled() As Boolean
            Get
                Return m_wildCardSupportEnabled
            End Get
        End Property
        Public Function GetPreFixList() As String()
            Dim i As Integer, hNav As System.Collections.IDictionaryEnumerator
            If Me.tokenObjects.Count > 0 Then
                hNav = tokenObjects.GetEnumerator()
                hNav.Reset()
                ReDim Me.prefixList(Me.tokenObjects.Count)
                For i = Me.tokenObjects.Count - 1 To 0 Step -1
                    hNav.MoveNext()
                    Me.prefixList(i) = hNav.Key
                Next
            End If
            Array.Sort(prefixList)
            Array.Reverse(prefixList)
            Return Me.prefixList
        End Function

        Public Function LoadMapFile(ByVal fName As String) As Boolean
            ' Create an instance of StreamReader to read from a file.
            Dim sr As StreamReader, fTXT As String
            sr = New StreamReader(fName)
            If File.Exists(fName) Then
                fTXT = sr.ReadToEnd
                sr.Close()
            Else
                Console.Error.WriteLine("File Not Found:" & fName)
                Return False
            End If

            fTXT = fTXT.Replace(Chr(13), "")
            Me.tokenList = fTXT.Split(Chr(10))
            Me.ProcessTokenList()

            Return True
        End Function

        Public Function LoadMap(ByVal MapText As String) As Boolean
            ' Create an instance of StreamReader to read from a file.
            Dim fTXT As String
            fTXT = MapText.Replace(Chr(13), "")
            Me.tokenList = fTXT.Split(Chr(10))
            Me.ProcessTokenList()
            Return True
        End Function


        Sub ProcessTokenList()
            Dim i As Integer, tCnt As Integer, prefixName As String, openToken As String, closeToken As String, tIO As String
            tCnt = 0 : closeToken = "" : tIO = ""
            For i = 0 To Me.tokenList.Length - 1 Step 3
                If Len(Me.tokenList(i)) > 0 Then
                    tCnt = i

                    prefixName = Me.tokenList(tCnt).Substring(0, Me.tokenList(tCnt).IndexOf("="))
                    openToken = Me.tokenList(tCnt).Substring(Me.tokenList(tCnt).IndexOf("=") + 1)

                    tCnt += 1
                    If prefixName = Me.tokenList(tCnt).Substring(0, Me.tokenList(tCnt).IndexOf("=")) Then
                        closeToken = Me.tokenList(tCnt).Substring(Me.tokenList(tCnt).IndexOf("=") + 1)
                    End If

                    tCnt += 1
                    If prefixName = Me.tokenList(tCnt).Substring(0, Me.tokenList(tCnt).IndexOf("=")) Then
                        tIO = Me.tokenList(tCnt).Substring(Me.tokenList(tCnt).IndexOf("=") + 1)
                    End If
                    If openToken.Length > 0 And closeToken.Length > 0 Then
                        If Not Me.SetTokenPair(prefixName, openToken, closeToken, tIO) Then
                            Console.Error.WriteLine("Could not create this Token Map:[" & prefixName & ":" & openToken & ":" & closeToken & ":" & tIO & "]")
                        End If
                    Else
                        Console.Error.WriteLine("Bad Token Map:[" & prefixName & ":" & openToken & ":" & closeToken & ":" & tIO & "]")
                    End If

                    If prefixName = Me.m_TokenType.XPATH_FIELD_TOKEN Then
                        Console.Error.WriteLine("Reserved Token Prefix in Map:[" & prefixName & ":" & openToken & ":" & closeToken & ":" & tIO & "]")
                        Console.Error.WriteLine("Warning: Wild Card support for empty xpath fields is disabled! See: XPATH_FIELD_TOKEN")
                        Me.m_wildCardSupportEnabled = False
                    End If

                    'Console.Error.WriteLine("---Token Map:[" & prefixName & ":" & openToken & ":" & closeToken & ":" & tIO & "]")
                    openToken = ""
                    closeToken = ""
                    tIO = ""
                End If
            Next
            ' make cur result tokenPair
            Dim tokPair As New ObjectParser.TokenPair(Me.m_TokenType)
            tokPair.prefixName = Me.m_TokenType.CURRENT_RESULT_TOKEN
            tokPair.tokenIO = "O"
            tokPair.MakeTokens(Me.m_TokenType.CURRENT_RESULT_TOKEN, Me.m_TokenType.CURRENT_RESULT_TOKEN)
            Me.SetTokenPair(tokPair)
        End Sub

        ' set token pairs by String Value
        Function SetTokenPair(ByVal prefixName As String, ByVal openToken As String, ByVal closeToken As String, ByVal tokenIO As String) As Boolean
            Dim tokPair As New ObjectParser.TokenPair(Me.m_TokenType)
            If prefixName = Me.m_TokenType.CURRENT_RESULT_TOKEN Then
                Me.mapHasErrors = True
                Console.Error.WriteLine("Warning: Reserved Prefix Detected:[" & prefixName & "]")
                Return False
            Else
                tokPair.prefixName = prefixName
                tokPair.tokenIO = tokenIO
                tokPair.MakeTokens(openToken, closeToken)
                Return Me.SetTokenPair(tokPair)
            End If
        End Function

        ' set token pair by Pairing Object, auto-managed
        Function SetTokenPair(ByVal tokenPair As ObjectParser.TokenPair) As Boolean
            If tokenPair.closeToken.hasCondition Then
                tokenPair.openToken.PrevNextTokenSize = tokenPair.closeToken.firstTagToken.Length
            End If
            If Me.tokenObjects.ContainsKey(tokenPair.prefixName) Then
                Me.tokenObjects.Item(tokenPair.prefixName) = tokenPair
            Else
                Me.tokenObjects.Add(tokenPair.prefixName, tokenPair)
            End If
            Return True
        End Function

        ' set token pair by Object, directly
        Function SetTokenPair(ByVal prefixName As String, ByVal openToken As ObjectParser.ParserToken, ByVal closeToken As ObjectParser.ParserToken, ByVal tokenIO As String) As Boolean
            Dim tokenPair As ObjectParser.TokenPair
            tokenPair = New ObjectParser.TokenPair(Me.m_TokenType)
            tokenPair.prefixName = prefixName

            openToken.SetOpenToken()
            closeToken.SetCloseToken()

            tokenPair.openToken = openToken
            tokenPair.closeToken = closeToken

            If closeToken.hasCondition Then
                openToken.PrevNextTokenSize = closeToken.firstTagToken.Length
                tokenPair.IO = tokenIO
            Else
                tokenPair.IO = tokenIO.Substring(0, 1)
            End If
            If Me.tokenObjects.ContainsKey(prefixName) Then
                If prefixName = Me.m_TokenType.CURRENT_RESULT_TOKEN Then
                    Me.mapHasErrors = True
                    Console.Error.WriteLine("Warning: Reserved Prefix Detected:[" & prefixName & "]")
                    Return False
                Else
                    Me.tokenObjects.Item(prefixName) = tokenPair
                End If
            Else
                Me.tokenObjects.Add(prefixName, tokenPair)
            End If
            Return True
        End Function

        ' get objects held in HashTable
        Function GetTokenPair(ByVal prefixName As String) As ObjectParser.TokenPair
            Dim tokenPair As ObjectParser.TokenPair, prefixName_ As String
            tokenPair = New ObjectParser.TokenPair(Me.m_TokenType)
            If prefixName Is Nothing Or prefixName Is Null Then
                prefixName_ = " "
            Else
                prefixName_ = prefixName
            End If
            ' the default prefix is one empty space, but may be presented as null or empty
            If prefixName_.Length = 0 Then
                prefixName_ = " "
            End If
            If Me.tokenObjects.ContainsKey(prefixName_) Then
                tokenPair = Me.tokenObjects.Item(prefixName_)
                tokenPair.openToken.Reset()
                tokenPair.openToken.SetOpenToken()
                tokenPair.closeToken.Reset()
                tokenPair.closeToken.SetCloseToken()
            End If
            Return tokenPair
        End Function

        ' see if objects are held in HashTable
        Function PrefixExists(ByVal prefixName As String) As Boolean
            Dim prefixName_ As String
            If prefixName Is Nothing Or prefixName Is Null Then
                prefixName_ = " "
            Else
                prefixName_ = prefixName
            End If
            ' the default prefix is one empty space, but may be presented as null or empty
            If prefixName_.Length = 0 Then
                prefixName_ = " "
            End If
            If Me.tokenObjects.ContainsKey(prefixName_) Then
                Return True
            End If
            Return False
        End Function

        Public Function getTokenMap() As String
            Dim tmpMap As StringBuilder, mapLine As String
            tmpMap = New StringBuilder
            For Each mapLine In Me.tokenList
                tmpMap.Append(mapLine)
                tmpMap.AppendLine()
            Next
            Return tmpMap.ToString()
        End Function

    End Class



End Namespace