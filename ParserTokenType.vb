
' -------------------------------------------------------------------------------------------
' Token Type Object, used by ParserToken Objects
' -------------------------------------------------------------------------------------------
Namespace ObjectParser

    Public Class ParserTokenType
        Public Const EMPTY As Integer = 16777216
        Public Const COMPLEX As Integer = 1
        Public Const JUMP As Integer = 2
        Public Const CONDITIONAL As Integer = 4

        Public Const SKIP As Integer = 8

        Public Const FORWARD As Integer = 16
        Public Const BACKWARD As Integer = 32

        Public Const XPATH_NODE As Integer = 64
        Public Const XPATH_ELEMENT As Integer = 128
        Public Const XPATH_ATTRIBUTE As Integer = 256
        Public Const XPATH_FILTER As Integer = 512
        Public Const XPATH_SUB As Integer = 1024
        Public Const XPATH_COL As Integer = 2048
        Public Const XPATH_LIST As Integer = 4096

        Public Const OPEN_TOKEN As Integer = 8192
        Public Const CLOSE_TOKEN As Integer = 16384
        Public Const NEXT_TOKEN As Integer = 32768

        Public Const INNER_RESULTS As Integer = 65536
        Public Const OUTER_RESULTS As Integer = 131072

        Public Const SEARCH_CONDITIONAL As Integer = 262144
        Public Const COMPARE_CONDITIONAL As Integer = 524288

        Public CURRENT_RESULT_TOKEN As String = "."
        Public JUMP_BACKWARD_TOKEN As String = "^^"
        Public JUMP_FORWARD_TOKEN As String = "^"
        Public CONDITIONAL_TOKEN As String = "?"
        Public CONDITIONAL_SPLIT As String = "|"
        Public QUOTE_TOKEN As String = "~"
        Public NAMESPACE_DELIM As String = ":"

        Public XPATH_SUB_OPEN As String = "{{"
        Public XPATH_SUB_CLOSE As String = "}}"

        Public XPATH_COL_OPEN As String = "("
        Public XPATH_COL_CLOSE As String = ")"

        Public XPATH_FILTER_OPEN As String = "["
        Public XPATH_FILTER_CLOSE As String = "]"

        Public XPATH_SPLIT As String = "/"

        Public XPATH_PREFIX_NODE As String = ""
        Public XPATH_PREFIX_ELEMENT As String = "!"
        Public XPATH_PREFIX_ATTRIBUTE As String = "@"

        Public XPATH_SUB_REF As String = "--SUB-PATH--"
        Public XPATH_FIELD_TOKEN As String = "*"
        Public XPATH_TOKEN_TEMP_DELIM As String = "|"

        Public MULTIPATH_OPEN As String = "{:"
        Public MULTIPATH_CLOSE As String = ":}"
        Public MULTIPATH_DELIM As String = ","
        Public MULTIPATH_PREV As String = "*"

        Public TypeMask As Integer

        Public Sub New()
            Me.TypeMask = ParserTokenType.EMPTY
        End Sub

        ' clone object
        Public Sub Clone(ByRef cloneTokenType As ObjectParser.ParserTokenType)

            cloneTokenType.CURRENT_RESULT_TOKEN = Me.CURRENT_RESULT_TOKEN
            cloneTokenType.JUMP_BACKWARD_TOKEN = Me.JUMP_BACKWARD_TOKEN
            cloneTokenType.JUMP_FORWARD_TOKEN = Me.JUMP_FORWARD_TOKEN
            cloneTokenType.CONDITIONAL_TOKEN = Me.CONDITIONAL_TOKEN
            cloneTokenType.CONDITIONAL_SPLIT = Me.CONDITIONAL_SPLIT
            cloneTokenType.QUOTE_TOKEN = Me.QUOTE_TOKEN
            cloneTokenType.NAMESPACE_DELIM = Me.NAMESPACE_DELIM

            cloneTokenType.XPATH_SUB_OPEN = Me.XPATH_SUB_OPEN
            cloneTokenType.XPATH_SUB_CLOSE = Me.XPATH_SUB_CLOSE
            cloneTokenType.XPATH_COL_OPEN = Me.XPATH_COL_OPEN
            cloneTokenType.XPATH_COL_CLOSE = Me.XPATH_COL_CLOSE
            cloneTokenType.XPATH_FILTER_OPEN = Me.XPATH_FILTER_OPEN
            cloneTokenType.XPATH_FILTER_CLOSE = Me.XPATH_FILTER_CLOSE
            cloneTokenType.XPATH_SPLIT = Me.XPATH_SPLIT
            cloneTokenType.XPATH_PREFIX_NODE = Me.XPATH_PREFIX_NODE
            cloneTokenType.XPATH_PREFIX_ELEMENT = Me.XPATH_PREFIX_ELEMENT
            cloneTokenType.XPATH_PREFIX_ATTRIBUTE = Me.XPATH_PREFIX_ATTRIBUTE
            cloneTokenType.XPATH_SUB_REF = Me.XPATH_SUB_REF
            cloneTokenType.XPATH_FIELD_TOKEN = Me.XPATH_FIELD_TOKEN
            cloneTokenType.XPATH_TOKEN_TEMP_DELIM = Me.XPATH_TOKEN_TEMP_DELIM

            cloneTokenType.MULTIPATH_OPEN = Me.MULTIPATH_OPEN
            cloneTokenType.MULTIPATH_CLOSE = Me.MULTIPATH_CLOSE
            cloneTokenType.MULTIPATH_DELIM = Me.MULTIPATH_DELIM
            cloneTokenType.MULTIPATH_PREV = Me.MULTIPATH_PREV

            cloneTokenType.TypeMask = ParserTokenType.EMPTY
        End Sub




















        Public Sub Init()
            Me.TypeMask = ParserTokenType.EMPTY
        End Sub

        Public Sub SetEmpty()
            Me.TypeMask = ParserTokenType.EMPTY
        End Sub

        Public Sub SetComplex()
            If Not Me.isComplex Then
                Me.TypeMask = Me.TypeMask Xor ParserTokenType.COMPLEX
            End If
        End Sub
        Public Sub SetConditional()
            If Not Me.hasCondition Then
                Me.TypeMask = Me.TypeMask Xor ParserTokenType.CONDITIONAL
            End If
        End Sub
        Public Sub SetJump()
            If Not Me.hasJump Then
                Me.TypeMask = Me.TypeMask Xor ParserTokenType.JUMP
            End If
        End Sub
        Public Sub SetSkip()
            If Not ParserTokenType.SKIP Then
                Me.TypeMask = Me.TypeMask Xor ParserTokenType.SKIP
            End If
        End Sub
        Public Sub SetForward()
            If Me.seekBackward Then
                Me.TypeMask = Me.TypeMask Xor ParserTokenType.BACKWARD
            End If
            If Not Me.seekForward Then
                Me.TypeMask = Me.TypeMask Xor ParserTokenType.FORWARD
            End If
        End Sub
        Public Sub SetBackward()
            If Me.seekForward Then
                Me.TypeMask = Me.TypeMask Xor ParserTokenType.FORWARD
            End If
            If Not Me.seekBackward Then
                Me.TypeMask = Me.TypeMask Xor ParserTokenType.BACKWARD
            End If
        End Sub

        Public Sub SetResultsInner()
            If Me.isOuterResults Then
                Me.TypeMask = Me.TypeMask Xor ParserTokenType.OUTER_RESULTS
            End If
            If Not Me.isInnerResults Then
                Me.TypeMask = Me.TypeMask Xor ParserTokenType.INNER_RESULTS
            End If
        End Sub
        Public Sub SetResultsOuter()
            If Me.isInnerResults Then
                Me.TypeMask = Me.TypeMask Xor ParserTokenType.INNER_RESULTS
            End If
            If Not Me.isOuterResults Then
                Me.TypeMask = Me.TypeMask Xor ParserTokenType.OUTER_RESULTS
            End If
        End Sub



        Public Sub SetOpenToken()
            If Me.isCloseToken Then
                Me.TypeMask = Me.TypeMask Xor ParserTokenType.CLOSE_TOKEN
            End If
            If Not Me.isOpenToken Then
                Me.TypeMask = Me.TypeMask Xor ParserTokenType.OPEN_TOKEN
            End If
        End Sub
        Public Sub SetCloseToken()
            If Me.isOpenToken Then
                Me.TypeMask = Me.TypeMask Xor ParserTokenType.OPEN_TOKEN
            End If
            If Not Me.isCloseToken Then
                Me.TypeMask = Me.TypeMask Xor ParserTokenType.CLOSE_TOKEN
            End If
        End Sub
        Public Sub SetNextToken()
            If Not Me.isNextToken Then
                Me.TypeMask = Me.TypeMask Xor ParserTokenType.NEXT_TOKEN
            End If
        End Sub

        Public Sub SetSearchConditional()
            If Not Me.isConditionalSearched Then
                Me.TypeMask = Me.TypeMask Xor ParserTokenType.SEARCH_CONDITIONAL
            End If
            If Me.isConditionalCompared Then
                Me.TypeMask = Me.TypeMask Xor ParserTokenType.COMPARE_CONDITIONAL
            End If
        End Sub
        Public Sub SetCompareConditional()
            If Not Me.isConditionalCompared Then
                Me.TypeMask = Me.TypeMask Xor ParserTokenType.COMPARE_CONDITIONAL
            End If
            If Me.isConditionalSearched Then
                Me.TypeMask = Me.TypeMask Xor ParserTokenType.SEARCH_CONDITIONAL
            End If
        End Sub



        Public Function isEmpty() As Boolean
            Return (Me.TypeMask = ParserTokenType.EMPTY)
        End Function

        Public Function isComplex() As Boolean
            Return (Me.TypeMask And ParserTokenType.COMPLEX) = ParserTokenType.COMPLEX
        End Function

        Public Function seekForward() As Boolean
            Return (Me.TypeMask And ParserTokenType.FORWARD) = ParserTokenType.FORWARD
        End Function
        Public Function seekBackward() As Boolean
            Return (Me.TypeMask And ParserTokenType.BACKWARD) = ParserTokenType.BACKWARD
        End Function


        Public Function hasCondition() As Boolean
            Return (Me.TypeMask And ParserTokenType.CONDITIONAL) = ParserTokenType.CONDITIONAL
        End Function
        Public Function hasJump() As Boolean
            Return (Me.TypeMask And ParserTokenType.JUMP) = ParserTokenType.JUMP
        End Function
        Public Function hasSkip() As Boolean
            Return (Me.TypeMask And ParserTokenType.SKIP) = ParserTokenType.SKIP
        End Function

        Public Function isOpenToken() As Boolean
            Return (Me.TypeMask And ParserTokenType.OPEN_TOKEN) = ParserTokenType.OPEN_TOKEN
        End Function
        Public Function isCloseToken() As Boolean
            Return (Me.TypeMask And ParserTokenType.CLOSE_TOKEN) = ParserTokenType.CLOSE_TOKEN
        End Function
        Public Function isNextToken() As Boolean
            Return (Me.TypeMask And ParserTokenType.NEXT_TOKEN) = ParserTokenType.NEXT_TOKEN
        End Function

        Public Function isInnerResults() As Boolean
            Return (Me.TypeMask And ParserTokenType.INNER_RESULTS) = ParserTokenType.INNER_RESULTS
        End Function
        Public Function isOuterResults() As Boolean
            Return (Me.TypeMask And ParserTokenType.OUTER_RESULTS) = ParserTokenType.OUTER_RESULTS
        End Function

        Public Function isConditionalSearched() As Boolean
            Return (Me.TypeMask And ParserTokenType.SEARCH_CONDITIONAL) = ParserTokenType.SEARCH_CONDITIONAL
        End Function
        Public Function isConditionalCompared() As Boolean
            Return (Me.TypeMask And ParserTokenType.COMPARE_CONDITIONAL) = ParserTokenType.COMPARE_CONDITIONAL
        End Function

    End Class

End Namespace


