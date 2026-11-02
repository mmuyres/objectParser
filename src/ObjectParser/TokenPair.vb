Namespace ObjectParser

    ''' <summary>One token map entry: a prefix with its open and close tokens and result scope.</summary>
    Public NotInheritable Class TokenDefinition
        Public Sub New(prefix As String, openToken As String, closeToken As String, ioSpec As String)
            Me.Prefix = prefix
            Me.OpenToken = openToken
            Me.CloseToken = closeToken
            Me.IOSpec = If(String.IsNullOrEmpty(ioSpec), "I", ioSpec)
        End Sub

        Public ReadOnly Property Prefix As String
        Public ReadOnly Property OpenToken As String
        Public ReadOnly Property CloseToken As String
        ''' <summary><c>I</c> (inner) or <c>O</c> (outer), or <c>X|Y</c> per branch of a conditional close token.</summary>
        Public ReadOnly Property IOSpec As String

        Public Overrides Function ToString() As String
            Return $"[{Prefix}] {OpenToken} .. {CloseToken} ({IOSpec})"
        End Function
    End Class

    ''' <summary>
    ''' A live open/close token pair used for one search. Pairs are created fresh from a
    ''' <see cref="TokenDefinition"/> for every path segment, so nested parsers (filters,
    ''' sub-paths) never disturb each other's match state.
    ''' </summary>
    Public NotInheritable Class TokenPair

        Friend Sub New(syntax As ParserSyntax, definition As TokenDefinition)
            Me.Definition = definition
            IsCurrentResult = definition.Prefix = syntax.CurrentResultToken

            OpenToken = New ParserToken(syntax, definition.OpenToken, TokenRole.Open, Me)
            CloseToken = New ParserToken(syntax, definition.CloseToken, TokenRole.Close, Me)

            Dim io = definition.IOSpec
            If CloseToken.HasCondition AndAlso io.Length < 3 Then
                io = io & syntax.ConditionalSplit & io
            ElseIf Not CloseToken.HasCondition AndAlso io.Length > 1 Then
                io = io.Substring(0, 1)
            End If
            IOSpec = io
            Me.IO = io

            If CloseToken.HasCondition Then
                OpenToken.LookaheadSize = Math.Max(1, CloseToken.FirstTagToken.Length)
                OpenToken.Reset()
            End If
        End Sub

        Public ReadOnly Property Definition As TokenDefinition
        Public ReadOnly Property OpenToken As ParserToken
        Public ReadOnly Property CloseToken As ParserToken
        ''' <summary>The normalised I/O spec (one entry per close-token branch).</summary>
        Public ReadOnly Property IOSpec As String
        ''' <summary>The scope chosen for the most recent match: <c>I</c> or <c>O</c>.</summary>
        Public Property IO As String
        ''' <summary>True for the built-in <c>.</c> pair that passes results through.</summary>
        Public ReadOnly Property IsCurrentResult As Boolean
        ''' <summary>Set once the parser derived the field name from the text (empty <c>*</c> field).</summary>
        Friend Property IsWildCarding As Boolean

        Public ReadOnly Property Prefix As String
            Get
                Return Definition.Prefix
            End Get
        End Property

        Friend Sub Reset()
            OpenToken.Reset()
            CloseToken.Reset()
        End Sub

    End Class

End Namespace
