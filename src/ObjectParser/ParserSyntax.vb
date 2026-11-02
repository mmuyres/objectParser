Namespace ObjectParser

    ''' <summary>
    ''' The special strings used by token maps and paths. One instance is shared by a
    ''' <see cref="TokenMap"/> and every <see cref="Parser"/> built on it, so change values
    ''' before building the map.
    ''' </summary>
    Public NotInheritable Class ParserSyntax

        ''' <summary>Path segment that passes the current result through unchanged (e.g. <c>//a[x]/.[y]</c>).</summary>
        Public Property CurrentResultToken As String = "."
        ''' <summary>In a token: find the text on the left, then jump backward-search to the text on the right.</summary>
        Public Property JumpBackwardToken As String = "^^"
        ''' <summary>In a token: find the text on the left, then jump forward to the text on the right.</summary>
        Public Property JumpForwardToken As String = "^"
        ''' <summary>In a close token: <c>a?b|c</c> uses b when the open match ended in a, otherwise c.</summary>
        Public Property ConditionalToken As String = "?"
        ''' <summary>Separates alternatives in conditional tokens and in a pair's I/O spec.</summary>
        Public Property ConditionalSplit As String = "|"
        ''' <summary>Stands for a double quote in manual delimiters (<c>-d</c>/<c>-D</c>).</summary>
        Public Property QuoteToken As String = "~"
        ''' <summary>
        ''' Starts a balanced close token: <c>%{}</c> matches through the brace that closes the next
        ''' <c>{</c>; <c>%{};</c> also fails if a <c>;</c> comes before that <c>{</c>.
        ''' </summary>
        Public Property BalancedToken As String = "%"

        Public Property SubPathOpen As String = "{{"
        Public Property SubPathClose As String = "}}"
        Public Property FilterOpen As String = "["
        Public Property FilterClose As String = "]"
        Public Property PathSeparator As String = "/"

        ''' <summary>Placeholder in a token that is replaced by the path segment's field name.</summary>
        Public Property FieldToken As String = "*"

        Public Property MultiPathOpen As String = "{:"
        Public Property MultiPathClose As String = ":}"
        Public Property MultiPathDelimiter As String = ","
        ''' <summary><c>{:*:}</c> re-uses the previous multi-path list.</summary>
        Public Property MultiPathPrevious As String = "*"

        ''' <summary>The prefix used when a path segment has no recognised prefix.</summary>
        Public Const DefaultPrefix As String = " "

        Friend Const SubPathReference As String = "--SUB-PATH--"

    End Class

End Namespace
