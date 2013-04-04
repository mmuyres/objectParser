'Imports System.IO
'Imports System.Text
'Imports Microsoft.VisualBasic.Strings
'Imports System.IO.FileSystemInfo
'Imports System.IO.FileInfo
'Imports System.IO.StreamWriter


Namespace ObjectParser
    ' -------------------------------------------------------------------------------------------
    ' Token Object Used by Object Parser
    ' -------------------------------------------------------------------------------------------
    Public Class ParserToken
        Private parentObj As ObjectParser.TokenPair
        Private seekError As Boolean

        Private strToken As String
        Private strTokenReset As String
        'Public TokenIO As String

        Private tokenFieldVal As String
        Private tokenNameSpace As String

        Private OuterBeginPos As Integer
        Private InnerBeginPos As Integer

        Private firstTag As String
        Private nextTag As String
        Private nextTag_CondVal As String

        Private oTagPos As Integer
        Private nTagPos As Integer

        Public firstTagNextChars As String
        Public NextTagPrevChars As String
        Public NextPrevCharsSize As Integer

        Private m_TokenType As ObjectParser.ParserTokenType

        Public Sub New(ByVal tokenTypeDef As ObjectParser.ParserTokenType)
            Me.m_TokenType = New ObjectParser.ParserTokenType
            tokenTypeDef.Clone(Me.m_TokenType)
            Me.ClearAll()
            Me.strToken = ""
            Me.strTokenReset = ""
            '            Me.TokenIO = "O"
            Me.NextPrevCharsSize = 2
        End Sub
        Public Sub New(ByVal tokenTypeDef As ObjectParser.ParserTokenType, ByVal TokenString As String)
            Me.New(tokenTypeDef)
            Me.Init(TokenString)
        End Sub
        ' get a ref to the TokenPair for common signaling
        Public Sub RegisterParent(ByRef parObj As ObjectParser.TokenPair)
            Me.parentObj = parObj
        End Sub
        ' clone setting to another Token
        Public Sub Clone(ByRef cpyToken As ParserToken)
            If cpyToken Is Nothing Then
                cpyToken = New ParserToken(m_TokenType)
            End If
            cpyToken.Clear()
            cpyToken.parentObj = Me.parentObj
            cpyToken.seekError = Me.seekError

            cpyToken.strToken = Me.strToken
            cpyToken.strTokenReset = Me.strTokenReset

            cpyToken.tokenFieldVal = Me.tokenFieldVal
            cpyToken.tokenNameSpace = Me.tokenNameSpace

            cpyToken.OuterBeginPos = Me.OuterBeginPos
            cpyToken.InnerBeginPos = Me.InnerBeginPos

            cpyToken.firstTag = Me.firstTag
            cpyToken.nextTag = Me.nextTag
            cpyToken.nextTag_CondVal = Me.nextTag_CondVal

            cpyToken.oTagPos = Me.oTagPos
            cpyToken.nTagPos = Me.nTagPos

            cpyToken.firstTagNextChars = Me.firstTagNextChars
            cpyToken.NextTagPrevChars = Me.NextTagPrevChars
            cpyToken.NextPrevCharsSize = Me.NextPrevCharsSize

            cpyToken.m_TokenType = Me.m_TokenType
            cpyToken.m_TokenType.Init()
        End Sub
        Public Sub Clear()
            Me.firstTag = ""
            Me.nextTag = ""
            Me.nextTag_CondVal = ""
            Me.firstTagNextChars = ""
            Me.NextTagPrevChars = ""
            Me.NextPrevCharsSize = 2
            seekError = False
            Me.m_TokenType.SetEmpty()
        End Sub
        Public Sub ClearAll()
            Me.Clear()
            Me.m_TokenType.SetEmpty()
            Me.OuterBeginPos = 0
            Me.InnerBeginPos = 0
            Me.oTagPos = 0
            Me.nTagPos = 0
            tokenFieldVal = ""
            tokenNameSpace = ""
        End Sub
        Public Sub ClearErrors()
            seekError = False
        End Sub
        Public Sub ClearBeginPos()
            Me.OuterBeginPos = 0
            Me.InnerBeginPos = 0
            Me.oTagPos = 0
            Me.nTagPos = 0
            seekError = False
        End Sub

        Public Sub Init(ByVal TokenString As String)

            Dim tPos As Integer, sTemp As String
            Me.strToken = TokenString
            If Me.strTokenReset.Length = 0 Then
                Me.strTokenReset = TokenString
            End If

            ' checks should be done in complexity order
            Me.Clear()
            '            Me.m_TokenType.SetEmpty()
            ' check for conditional compare token
            If Me.m_TokenType.isEmpty() Then
                tPos = Me.strToken.IndexOf(Me.m_TokenType.CONDITIONAL_TOKEN)
                If tPos > -1 Then
                    Me.firstTag = Me.strToken.Substring(0, tPos)
                    sTemp = Me.strToken.Substring(tPos + Me.m_TokenType.CONDITIONAL_TOKEN.Length)

                    tPos = sTemp.IndexOf(Me.m_TokenType.CONDITIONAL_SPLIT)
                    If tPos > -1 Then
                        Me.m_TokenType.SetComplex()
                        Me.m_TokenType.SetConditional()
                        Me.m_TokenType.SetForward()
                        Me.m_TokenType.SetCompareConditional()
                        Me.nextTag = sTemp
                    Else
                        Me.firstTag = Me.strToken
                        Me.nextTag = ""
                    End If
                End If
            End If

            ' check for conditional search token
            If Me.m_TokenType.isEmpty() Then
                tPos = Me.strToken.IndexOf(Me.m_TokenType.CONDITIONAL_SPLIT)
                If tPos > -1 Then
                    Me.firstTag = Me.strToken.Substring(0, tPos)
                    sTemp = Me.strToken.Substring(tPos + Me.m_TokenType.CONDITIONAL_SPLIT.Length)
                    tPos = sTemp.IndexOf(Me.m_TokenType.CONDITIONAL_SPLIT)
                    If tPos > -1 Then
                        Me.m_TokenType.SetComplex()
                        Me.m_TokenType.SetConditional()
                        Me.m_TokenType.SetBackward()
                        Me.m_TokenType.SetSearchConditional()
                        Me.nextTag = sTemp
                    Else
                        Me.firstTag = Me.strToken
                        Me.nextTag = ""
                    End If
                End If
            End If

            ' check for Jump backwards token
            If Me.m_TokenType.isEmpty() Then
                tPos = Me.strToken.IndexOf(Me.m_TokenType.JUMP_BACKWARD_TOKEN)
                If tPos > -1 Then
                    Me.m_TokenType.SetComplex()
                    Me.m_TokenType.SetJump()
                    Me.m_TokenType.SetBackward()
                    Me.firstTag = Me.strToken.Substring(0, tPos)
                    Me.nextTag = Me.strToken.Substring(tPos + Me.m_TokenType.JUMP_BACKWARD_TOKEN.Length)
                    '                    Me.NextPrevCharsSize = Me.nextTag.Length
                End If
            End If

            ' check for Jump Forward Token
            If Me.m_TokenType.isEmpty() Then
                tPos = Me.strToken.IndexOf(Me.m_TokenType.JUMP_FORWARD_TOKEN)
                If tPos > -1 Then
                    Me.m_TokenType.SetJump()
                    Me.m_TokenType.SetForward()
                    Me.firstTag = Me.strToken.Substring(0, tPos)
                    Me.nextTag = Me.strToken.Substring(tPos + Me.m_TokenType.JUMP_FORWARD_TOKEN.Length)
                End If
            End If

            ' must be a simple token
            If Me.m_TokenType.isEmpty() Then
                Me.firstTag = Me.strToken
                Me.m_TokenType.SetForward()
                Me.nextTag = ""
            End If

        End Sub

        Public Sub Reset()
            '            Me.Clear()
            '           Me.ClearBeginPos()
            Me.Init(Me.strTokenReset)
        End Sub

        Public ReadOnly Property hadError() As Boolean
            Get
                Return seekError
            End Get
        End Property
        Public Property FieldValue() As String
            Get
                Return Me.tokenFieldVal
            End Get
            Set(ByVal Value As String)
                Me.tokenFieldVal = Value
            End Set
        End Property

        Public Property NameSpaceValue() As String
            Get
                Return Me.tokenNameSpace
            End Get
            Set(ByVal Value As String)
                Me.tokenNameSpace = Value
            End Set
        End Property
        Public Property PrevNextTokenSize() As Integer
            Get
                Return Me.NextPrevCharsSize
            End Get
            Set(ByVal Value As Integer)
                Me.NextPrevCharsSize = Value
            End Set
        End Property

        Public ReadOnly Property isComplex() As Boolean
            Get
                Return Me.m_TokenType.isComplex
            End Get
        End Property
        Public ReadOnly Property hasCondition() As Boolean
            Get
                Return Me.m_TokenType.hasCondition
            End Get
        End Property
        Public ReadOnly Property hasJump() As Boolean
            Get
                Return Me.m_TokenType.hasJump
            End Get
        End Property
        Public ReadOnly Property hasSkip() As Boolean
            Get
                Return Me.m_TokenType.hasSkip
            End Get
        End Property
        Public ReadOnly Property Skip() As Boolean
            Get
                Return Me.m_TokenType.hasSkip()
            End Get
        End Property

        Public ReadOnly Property isCompareConditional() As Boolean
            Get
                Return Me.m_TokenType.isConditionalCompared
            End Get
        End Property
        Public ReadOnly Property isSearchConditional() As Boolean
            Get
                Return Me.m_TokenType.isConditionalSearched
            End Get
        End Property


        Public Sub SetOpenToken()
            Me.m_TokenType.SetOpenToken()
        End Sub
        Public Sub SetCloseToken()
            Me.m_TokenType.SetCloseToken()
        End Sub
        Public Sub SetNextToken()
            Me.m_TokenType.SetNextToken()
        End Sub

        Public ReadOnly Property FieldIsEmpty() As Boolean
            Get
                Return (Me.tokenFieldVal.Length = 0)
            End Get
        End Property
        Public ReadOnly Property FieldIsWild() As Boolean
            Get
                Return (Me.firstTagHasFieldToken AndAlso (Me.tokenFieldVal.Length = 0 OrElse Me.tokenFieldVal = Me.m_TokenType.XPATH_FIELD_TOKEN))
            End Get
        End Property


        Public Property firstTagPos() As Integer
            Get
                Return Me.oTagPos
            End Get
            Set(ByVal Value As Integer)
                If Value > -1 Then
                    oTagPos = Value
                    If Me.m_TokenType.isOpenToken Then
                        Me.OuterBeginPos = Value
                        Me.InnerBeginPos = Value + Me.firstTagVal.Length
                    Else
                        Me.InnerBeginPos = Value
                        Me.OuterBeginPos = Value + Me.firstTagVal.Length
                    End If
                    seekError = False
                Else
                    seekError = True
                End If
            End Set
        End Property
        Public Property nextTagPos() As Integer
            Get
                Return Me.nTagPos
            End Get
            Set(ByVal Value As Integer)
                If Value > -1 Then
                    Me.nTagPos = Value
                    If Me.m_TokenType.isOpenToken Then
                        Me.InnerBeginPos = Value + Me.nextTagVal.Length
                    Else
                        Me.OuterBeginPos = Value + Me.nextTagVal.Length
                    End If
                    seekError = False
                Else
                    seekError = True
                End If
            End Set
        End Property


        Public ReadOnly Property firstPosEnd() As Integer
            Get
                Dim fPosEnd As Integer
                fPosEnd = Me.oTagPos + Me.firstTag.Replace(Me.m_TokenType.XPATH_FIELD_TOKEN, "").Length
                If Me.firstTagHasFieldToken Then
                    fPosEnd += Me.FieldValue.Trim.Length
                End If
                Return fPosEnd
            End Get
        End Property
        Public ReadOnly Property nextPosEnd() As Integer
            Get
                Return Me.nTagPos + Me.nextTagVal.Length
            End Get
        End Property


        ' Inner and Outer Positions
        Public Property OuterPos() As Integer
            Get
                Return Me.OuterBeginPos
            End Get
            Set(ByVal Value As Integer)
                Me.OuterBeginPos = Value
            End Set
        End Property
        Public Property InnerPos() As Integer
            Get
                Return Me.InnerBeginPos
            End Get
            Set(ByVal Value As Integer)
                Me.InnerBeginPos = Value
            End Set
        End Property
        Public ReadOnly Property firstTagToken() As String
            Get
                Return Me.firstTag
            End Get
        End Property
        Public ReadOnly Property firstTagHasFieldToken() As String
            Get
                Return (Me.firstTag.IndexOf(Me.m_TokenType.XPATH_FIELD_TOKEN) <> -1)
            End Get
        End Property

        Public ReadOnly Property firstTagVal() As String
            Get
                If Me.tokenFieldVal.Length > 0 AndAlso Me.m_TokenType.XPATH_FIELD_TOKEN <> Me.tokenFieldVal Then
                    Return Me.firstTag.Replace(Me.m_TokenType.XPATH_FIELD_TOKEN, Me.tokenFieldVal)
                Else
                    Return Me.firstTag.Replace(Me.m_TokenType.XPATH_FIELD_TOKEN, "")
                End If
            End Get
        End Property
        Public ReadOnly Property nextTagToken() As String
            Get
                Return Me.nextTag
            End Get
        End Property
        Public ReadOnly Property nextTagHasFieldToken() As String
            Get
                Return (Me.nextTag.IndexOf(Me.m_TokenType.XPATH_FIELD_TOKEN) <> -1)
            End Get
        End Property
        ' conditional value used for next tags conditional search
        Public Property nextTagCondVal() As String
            Get
                Return Me.nextTag_CondVal
            End Get
            Set(ByVal value As String)
                Me.nextTag_CondVal = value
            End Set
        End Property

        ' tag val may change if it is a conditional token
        Public ReadOnly Property nextTagVal() As String
            Get
                If Me.tokenFieldVal.Length > 0 AndAlso Me.m_TokenType.XPATH_FIELD_TOKEN <> Me.tokenFieldVal Then
                    If m_TokenType.isConditionalSearched Then
                        Return Me.nextTagCondVal.Replace(Me.m_TokenType.XPATH_FIELD_TOKEN, Me.tokenFieldVal)
                    Else
                        Return Me.nextTag.Replace(Me.m_TokenType.XPATH_FIELD_TOKEN, Me.tokenFieldVal)
                    End If
                Else
                    If m_TokenType.isConditionalSearched Then
                        Return Me.nextTagCondVal.Replace(Me.m_TokenType.XPATH_FIELD_TOKEN, "")
                    Else
                        Return Me.nextTag.Replace(Me.m_TokenType.XPATH_FIELD_TOKEN, "")
                    End If
                End If
            End Get
        End Property
        Public ReadOnly Property nextTagCond_A() As String
            Get
                If Me.m_TokenType.hasCondition Then
                    Return Me.nextTag.Split(Me.m_TokenType.CONDITIONAL_SPLIT)(0)
                Else
                    Return Me.nextTag
                End If
            End Get
        End Property
        Public ReadOnly Property nextTagCond_B() As String
            Get
                If Me.m_TokenType.hasCondition Then
                    Return Me.nextTag.Split(Me.m_TokenType.CONDITIONAL_SPLIT)(1)
                Else
                    Return Me.nextTag
                End If
            End Get
        End Property

        Public ReadOnly Property nextTagConditions() As String()
            Get
                If Me.m_TokenType.hasCondition Then
                    Return Me.nextTag.Split(Me.m_TokenType.CONDITIONAL_SPLIT, 256, System.StringSplitOptions.None)
                Else
                    Dim sList(1) As String
                    sList(0) = Me.nextTag
                    Return sList
                End If
            End Get
        End Property

        Public ReadOnly Property Parent() As ObjectParser.TokenPair
            Get
                Return Me.parentObj
            End Get
        End Property
        Property Type() As ObjectParser.ParserTokenType
            Get
                Return Me.m_TokenType
            End Get
            Set(ByVal Value As ObjectParser.ParserTokenType)
                Me.m_TokenType = New ObjectParser.ParserTokenType
                Value.Clone(m_TokenType)
            End Set
        End Property


        Public Function equalsNextChar(ByVal charValue As String) As Boolean
            If charValue.Length > 0 And Me.firstTagNextChars.Length > 0 And Me.firstTagNextChars = charValue Then
                Return True
            Else
                Return False
            End If
        End Function
        Public Function equalsPrevChar(ByVal charValue As String) As Boolean
            If charValue.Length > 0 And Me.NextTagPrevChars.Length > 0 And Me.NextTagPrevChars = charValue Then
                Return True
            Else
                Return False
            End If
        End Function

        ' normally called by the CloseToken to see if it has anything to do or needs to adjust its token
        Public Sub EvalCondition(ByRef firstToken As ObjectParser.ParserToken)
            Dim tmpTag As String, ioType As String
            ioType = "O"
            ' ensure we are a conditional comparative Token
            If Me.hasCondition AndAlso Me.isCompareConditional Then
                ' see if the  char after the firstTag or the Char before the Next Tag from the prev Token Object's is our first Tag, if so use the first conditional - else the second conditional for our Next Tag
                If firstToken.equalsNextChar(Me.firstTag) Or firstToken.equalsPrevChar(Me.firstTag) Then
                    tmpTag = Me.nextTagCond_A
                    ioType = Me.parentObj.tokenIO.Split(Me.m_TokenType.CONDITIONAL_SPLIT)(0)
                Else
                    tmpTag = Me.nextTagCond_B
                    ioType = Me.Parent.tokenIO.Split(Me.m_TokenType.CONDITIONAL_SPLIT)(1)
                End If
                Me.Init(tmpTag)
            ElseIf Me.hasCondition AndAlso Me.isSearchConditional Then
                ioType = Me.parentObj.tokenIO.Split(Me.m_TokenType.CONDITIONAL_SPLIT)(0)
            Else
                ioType = Me.parentObj.tokenIO
            End If

            ' adjust results scope preferred, should be put into TokenPair object
            Me.parentObj.IO = ioType

            If ioType = "I" Then
                Me.m_TokenType.SetResultsInner()
            Else
                Me.m_TokenType.SetResultsOuter()
            End If

            ' check to see if the other Token Object has already found our next token, if so take his end position and then reset is end position to his Inner Position
            If firstToken.equalsNextChar(Me.firstTag) Or firstToken.equalsPrevChar(Me.firstTag) Then
                If Me.Parent.debug Then
                    Console.Out.WriteLine("--First Tag Equal Next Tag: " & Me.firstTag)
                End If
                Me.m_TokenType.SetSkip()
                firstToken.InnerBeginPos = firstToken.firstPosEnd
                Me.oTagPos = firstToken.firstTagPos - firstToken.firstTagVal.Length
                Me.nTagPos = firstToken.nextPosEnd
                Me.InnerBeginPos = firstToken.nextTagPos - firstToken.nextTag.Length
                Me.OuterBeginPos = firstToken.nextPosEnd
            End If
        End Sub
    End Class

End Namespace