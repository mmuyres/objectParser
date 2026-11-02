Namespace ObjectParser

    ''' <summary>How a token string was classified when it was parsed.</summary>
    <Flags>
    Public Enum TokenFlags
        None = 0
        Complex = 1
        Jump = 2
        Conditional = 4
        Skip = 8
        Forward = 16
        Backward = 32
        InnerResults = 64
        OuterResults = 128
        SearchConditional = 256
        CompareConditional = 512
        Balanced = 1024
        FieldResults = 2048
    End Enum

    ''' <summary>Whether a token opens or closes a match.</summary>
    Public Enum TokenRole
        Open
        Close
    End Enum

    ''' <summary>
    ''' One half of a <see cref="TokenPair"/>: a delimiter definition such as <c>&lt;*^&gt;</c>
    ''' plus the positions it matched during the current search.
    ''' </summary>
    ''' <remarks>
    ''' Token grammar (checked in this order):
    ''' <list type="bullet">
    ''' <item><c>a?b|c</c> compare-conditional: if the open match ended in <c>a</c> use <c>b</c>, else <c>c</c>.</item>
    ''' <item><c>a|b|c</c> search-conditional: find <c>a</c>, then the nearest of the alternatives.</item>
    ''' <item><c>a^^b</c> jump backward, <c>a^b</c> jump forward: find <c>a</c>, then <c>b</c>.</item>
    ''' <item>anything else is a literal. <c>*</c> is replaced by the path field name.</item>
    ''' </list>
    ''' </remarks>
    Public NotInheritable Class ParserToken

        Private ReadOnly _syntax As ParserSyntax
        Private ReadOnly _definition As String
        Private _flags As TokenFlags

        Private _firstTag As String = ""
        Private _nextTag As String = ""

        Friend Sub New(syntax As ParserSyntax, definition As String, role As TokenRole, parent As TokenPair)
            _syntax = syntax
            _definition = If(definition, "")
            Me.Role = role
            Me.Parent = parent
            FieldValue = ""
            Init(_definition)
        End Sub

        Public ReadOnly Property Role As TokenRole
        Public ReadOnly Property Parent As TokenPair
        Public ReadOnly Property Definition As String
            Get
                Return _definition
            End Get
        End Property
        Public ReadOnly Property Flags As TokenFlags
            Get
                Return _flags
            End Get
        End Property

        ''' <summary>The field name substituted for <c>*</c>.</summary>
        Public Property FieldValue As String

        Friend Property HadError As Boolean
        Friend Property OuterPos As Integer
        Friend Property InnerPos As Integer
        Friend Property NextTagConditionValue As String = ""
        Friend Property FirstTagNextChars As String = ""
        Friend Property NextTagPrevChars As String = ""
        Friend Property NextPrevCharsSize As Integer = 2

        ''' <summary>
        ''' How many characters after the first tag are remembered so a compare-conditional close
        ''' token can test them (the length of the close token's condition, e.g. 2 for <c>/&gt;</c>).
        ''' </summary>
        Friend Property LookaheadSize As Integer = 2

        Private _firstTagPos As Integer
        Private _nextTagPos As Integer

#Region "Parsing the definition"

        Friend Sub Init(tokenText As String)
            _flags = TokenFlags.None
            _firstTag = ""
            _nextTag = ""
            NextTagConditionValue = ""
            FirstTagNextChars = ""
            NextTagPrevChars = ""
            NextPrevCharsSize = LookaheadSize
            HadError = False

            Dim pos As Integer

            ' balanced: %{}  or %{}; (open char, close char, optional stop char)
            Dim marker = _syntax.BalancedToken
            If marker.Length > 0 AndAlso tokenText.StartsWith(marker, StringComparison.Ordinal) AndAlso
               (tokenText.Length = marker.Length + 2 OrElse tokenText.Length = marker.Length + 3) Then
                BalancedOpen = tokenText(marker.Length)
                BalancedClose = tokenText(marker.Length + 1)
                BalancedStop = If(tokenText.Length = marker.Length + 3, tokenText(marker.Length + 2), ChrW(0))
                _flags = TokenFlags.Balanced Or TokenFlags.Forward
                Return
            End If

            ' compare-conditional: a?b|c
            pos = tokenText.IndexOf(_syntax.ConditionalToken, StringComparison.Ordinal)
            If pos > -1 Then
                Dim rest = tokenText.Substring(pos + _syntax.ConditionalToken.Length)
                If rest.IndexOf(_syntax.ConditionalSplit, StringComparison.Ordinal) > -1 Then
                    _firstTag = tokenText.Substring(0, pos)
                    _nextTag = rest
                    _flags = TokenFlags.Complex Or TokenFlags.Conditional Or TokenFlags.Forward Or TokenFlags.CompareConditional
                    Return
                End If
            End If

            ' search-conditional: a|b|c (needs at least two alternatives after a)
            pos = tokenText.IndexOf(_syntax.ConditionalSplit, StringComparison.Ordinal)
            If pos > -1 Then
                Dim rest = tokenText.Substring(pos + _syntax.ConditionalSplit.Length)
                If rest.IndexOf(_syntax.ConditionalSplit, StringComparison.Ordinal) > -1 Then
                    _firstTag = tokenText.Substring(0, pos)
                    _nextTag = rest
                    _flags = TokenFlags.Complex Or TokenFlags.Conditional Or TokenFlags.Backward Or TokenFlags.SearchConditional
                    Return
                End If
            End If

            ' jump backward: a^^b
            pos = tokenText.IndexOf(_syntax.JumpBackwardToken, StringComparison.Ordinal)
            If pos > -1 Then
                _firstTag = tokenText.Substring(0, pos)
                _nextTag = tokenText.Substring(pos + _syntax.JumpBackwardToken.Length)
                _flags = TokenFlags.Complex Or TokenFlags.Jump Or TokenFlags.Backward
                Return
            End If

            ' jump forward: a^b
            pos = tokenText.IndexOf(_syntax.JumpForwardToken, StringComparison.Ordinal)
            If pos > -1 Then
                _firstTag = tokenText.Substring(0, pos)
                _nextTag = tokenText.Substring(pos + _syntax.JumpForwardToken.Length)
                _flags = TokenFlags.Jump Or TokenFlags.Forward
                Return
            End If

            ' simple literal
            _firstTag = tokenText
            _flags = TokenFlags.Forward
        End Sub

        ''' <summary>Re-parses the original definition, dropping per-match state.</summary>
        Friend Sub Reset()
            Init(_definition)
        End Sub

        Friend Sub ClearPositions()
            OuterPos = 0
            InnerPos = 0
            _firstTagPos = 0
            _nextTagPos = 0
            HadError = False
        End Sub

#End Region

#Region "Flags"

        Private Function Has(flag As TokenFlags) As Boolean
            Return (_flags And flag) = flag
        End Function

        Public ReadOnly Property IsComplex As Boolean
            Get
                Return Has(TokenFlags.Complex)
            End Get
        End Property
        Public ReadOnly Property HasCondition As Boolean
            Get
                Return Has(TokenFlags.Conditional)
            End Get
        End Property
        Public ReadOnly Property HasJump As Boolean
            Get
                Return Has(TokenFlags.Jump)
            End Get
        End Property
        Public ReadOnly Property IsSearchConditional As Boolean
            Get
                Return Has(TokenFlags.SearchConditional)
            End Get
        End Property
        Public ReadOnly Property IsCompareConditional As Boolean
            Get
                Return Has(TokenFlags.CompareConditional)
            End Get
        End Property
        Friend ReadOnly Property Skip As Boolean
            Get
                Return Has(TokenFlags.Skip)
            End Get
        End Property
        Friend ReadOnly Property SeekBackward As Boolean
            Get
                Return Has(TokenFlags.Backward)
            End Get
        End Property
        Friend ReadOnly Property SeekForward As Boolean
            Get
                Return Has(TokenFlags.Forward)
            End Get
        End Property
        Friend ReadOnly Property IsOuterResults As Boolean
            Get
                Return Has(TokenFlags.OuterResults)
            End Get
        End Property
        ''' <summary>The pair's scope is <c>F</c>: the result is the field name read from the text.</summary>
        Friend ReadOnly Property IsFieldResults As Boolean
            Get
                Return Has(TokenFlags.FieldResults)
            End Get
        End Property
        Public ReadOnly Property IsBalanced As Boolean
            Get
                Return Has(TokenFlags.Balanced)
            End Get
        End Property
        Friend Property BalancedOpen As Char
        Friend Property BalancedClose As Char
        ''' <summary>ChrW(0) when the balanced token has no stop character.</summary>
        Friend Property BalancedStop As Char

        Private Sub SetResultScope(io As String)
            _flags = _flags And Not (TokenFlags.InnerResults Or TokenFlags.OuterResults Or TokenFlags.FieldResults)
            Select Case io
                Case "I" : _flags = _flags Or TokenFlags.InnerResults
                Case "F" : _flags = _flags Or TokenFlags.FieldResults
                Case Else : _flags = _flags Or TokenFlags.OuterResults
            End Select
        End Sub

#End Region

#Region "Tag values"

        Public ReadOnly Property FirstTagToken As String
            Get
                Return _firstTag
            End Get
        End Property
        Public ReadOnly Property NextTagToken As String
            Get
                Return _nextTag
            End Get
        End Property

        Friend ReadOnly Property FirstTagHasFieldToken As Boolean
            Get
                Return _firstTag.Contains(_syntax.FieldToken, StringComparison.Ordinal)
            End Get
        End Property

        Private ReadOnly Property HasConcreteField As Boolean
            Get
                Return FieldValue.Length > 0 AndAlso FieldValue <> _syntax.FieldToken
            End Get
        End Property

        Friend ReadOnly Property FieldIsWild As Boolean
            Get
                Return FirstTagHasFieldToken AndAlso Not HasConcreteField
            End Get
        End Property

        ''' <summary>The first tag with the field substituted.</summary>
        Friend ReadOnly Property FirstTagValue As String
            Get
                Return TokenMap.Literal(_firstTag.Replace(_syntax.FieldToken, If(HasConcreteField, FieldValue, ""), StringComparison.Ordinal))
            End Get
        End Property

        ''' <summary>The next tag with the field substituted; for search-conditionals, the alternative that matched.</summary>
        Friend ReadOnly Property NextTagValue As String
            Get
                ' NextTagConditionValue is already literal (it came from NextTagConditions)
                If IsSearchConditional Then
                    Return NextTagConditionValue.Replace(_syntax.FieldToken, If(HasConcreteField, FieldValue, ""), StringComparison.Ordinal)
                End If
                Return TokenMap.Literal(_nextTag.Replace(_syntax.FieldToken, If(HasConcreteField, FieldValue, ""), StringComparison.Ordinal))
            End Get
        End Property

        Friend ReadOnly Property NextTagConditions As String()
            Get
                If HasCondition Then
                    Return _nextTag.Split(_syntax.ConditionalSplit).Select(AddressOf TokenMap.Literal).ToArray()
                End If
                Return {TokenMap.Literal(_nextTag)}
            End Get
        End Property

        ''' <summary>The literal text before and after the field token in the first tag.</summary>
        Friend Sub GetFieldContext(ByRef lead As String, ByRef trail As String)
            Dim star = _firstTag.IndexOf(_syntax.FieldToken, StringComparison.Ordinal)
            lead = TokenMap.Literal(_firstTag.Substring(0, star))
            trail = TokenMap.Literal(_firstTag.Substring(star + _syntax.FieldToken.Length))
        End Sub

        ''' <summary>
        ''' True when the field sits at the very start or end of the first tag, so a match must not
        ''' run into neighbouring name characters (<c>&lt;choice</c> must not match <c>&lt;choices</c>).
        ''' </summary>
        Friend ReadOnly Property FieldNeedsTrailingBoundary As Boolean
            Get
                Return HasConcreteField AndAlso _firstTag.EndsWith(_syntax.FieldToken, StringComparison.Ordinal)
            End Get
        End Property
        Friend ReadOnly Property FieldNeedsLeadingBoundary As Boolean
            Get
                Return HasConcreteField AndAlso _firstTag.StartsWith(_syntax.FieldToken, StringComparison.Ordinal)
            End Get
        End Property

#End Region

#Region "Positions"

        Friend Property FirstTagPos As Integer
            Get
                Return _firstTagPos
            End Get
            Set(value As Integer)
                If value > -1 Then
                    _firstTagPos = value
                    If Role = TokenRole.Open Then
                        OuterPos = value
                        InnerPos = value + FirstTagValue.Length
                    Else
                        InnerPos = value
                        OuterPos = value + FirstTagValue.Length
                    End If
                    HadError = False
                Else
                    HadError = True
                End If
            End Set
        End Property

        Friend Property NextTagPos As Integer
            Get
                Return _nextTagPos
            End Get
            Set(value As Integer)
                If value > -1 Then
                    _nextTagPos = value
                    If Role = TokenRole.Open Then
                        InnerPos = value + NextTagValue.Length
                    Else
                        OuterPos = value + NextTagValue.Length
                        ' a close token with no first tag ("|a|b", "^a") is a lookahead: the match
                        ' ends where it found its target, and the target is left for the next match
                        If IsLookahead Then
                            InnerPos = value
                            OuterPos = value
                        End If
                    End If
                    HadError = False
                Else
                    HadError = True
                End If
            End Set
        End Property

        Friend ReadOnly Property FirstPosEnd As Integer
            Get
                Dim length = _firstTag.Replace(_syntax.FieldToken, "", StringComparison.Ordinal).Length
                If FirstTagHasFieldToken Then
                    length += FieldValue.Trim().Length
                End If
                Return _firstTagPos + length
            End Get
        End Property

        Friend ReadOnly Property NextPosEnd As Integer
            Get
                Return _nextTagPos + NextTagValue.Length
            End Get
        End Property

        ''' <summary>A close token with an empty first tag: it marks where a match ends without consuming it.</summary>
        Friend ReadOnly Property IsLookahead As Boolean
            Get
                Return Role = TokenRole.Close AndAlso _firstTag.Length = 0 AndAlso Not IsBalanced
            End Get
        End Property

        ''' <summary>Where the field text starts in the matched first tag.</summary>
        Friend ReadOnly Property FieldStart As Integer
            Get
                Dim star = _firstTag.IndexOf(_syntax.FieldToken, StringComparison.Ordinal)
                Return _firstTagPos + Math.Max(0, star)
            End Get
        End Property

        ''' <summary>Records a balanced match from the opening character at <paramref name="openAt"/> to the closing one at <paramref name="closeAt"/>.</summary>
        Friend Sub SetBalancedMatch(openAt As Integer, closeAt As Integer)
            _firstTagPos = closeAt
            _nextTagPos = closeAt
            InnerPos = closeAt
            OuterPos = closeAt + 1
            HadError = False
            If Role = TokenRole.Close Then Parent.OpenToken.InnerPos = openAt + 1
        End Sub

#End Region

#Region "Conditional evaluation"

        Private Function FollowsOrPrecedes(other As ParserToken, value As String) As Boolean
            value = TokenMap.Literal(value)
            If value.Length = 0 Then Return False
            Return (other.FirstTagNextChars.Length > 0 AndAlso other.FirstTagNextChars = value) OrElse
                   (other.NextTagPrevChars.Length > 0 AndAlso other.NextTagPrevChars = value)
        End Function

        ''' <summary>
        ''' Called on the close token after the open token matched. Picks the conditional branch,
        ''' sets inner/outer result scope, and marks the close token as already satisfied when the
        ''' open match already consumed it (e.g. a self-closing <c>/&gt;</c>).
        ''' </summary>
        Friend Sub EvalCondition(openToken As ParserToken)
            Dim ioSpec = Parent.IOSpec
            Dim io As String
            If HasCondition AndAlso IsCompareConditional Then
                Dim parts = ioSpec.Split(_syntax.ConditionalSplit)
                Dim conditions = NextTagConditions
                If FollowsOrPrecedes(openToken, _firstTag) Then
                    io = parts(0)
                    Init(conditions(0))
                Else
                    io = If(parts.Length > 1, parts(1), parts(0))
                    Init(If(conditions.Length > 1, conditions(1), conditions(0)))
                End If
            ElseIf HasCondition AndAlso IsSearchConditional Then
                io = ioSpec.Split(_syntax.ConditionalSplit)(0)
            Else
                io = ioSpec
            End If

            Parent.IO = io
            SetResultScope(io)

            If FollowsOrPrecedes(openToken, _firstTag) Then
                _flags = _flags Or TokenFlags.Skip
                openToken.InnerPos = openToken.FirstPosEnd
                _firstTagPos = openToken.FirstTagPos - openToken.FirstTagValue.Length
                _nextTagPos = openToken.NextPosEnd
                InnerPos = openToken.NextTagPos - openToken.NextTagToken.Length
                OuterPos = openToken.NextPosEnd
            End If
        End Sub

#End Region

        Public Overrides Function ToString() As String
            Return $"{Role}:{_definition}"
        End Function

    End Class

End Namespace
