Imports System.IO
Imports System.Text

Namespace ObjectParser

    ''' <summary>
    ''' The set of token pairs a <see cref="Parser"/> understands, keyed by path prefix.
    ''' </summary>
    ''' <remarks>
    ''' Map text is a series of three-line entries for each prefix:
    ''' <code>
    ''' prefix=&lt;open token&gt;
    ''' prefix=&lt;close token&gt;
    ''' prefix=I|O
    ''' </code>
    ''' The default prefix is a single space. Blank lines between entries are ignored.
    ''' </remarks>
    Public NotInheritable Class TokenMap

        Private ReadOnly _definitions As New Dictionary(Of String, TokenDefinition)(StringComparer.Ordinal)
        Private ReadOnly _warnings As New List(Of String)
        Private _prefixes As String()
        Private _sourceText As String = ""

        Public Sub New()
            Me.New(New ParserSyntax())
        End Sub

        Public Sub New(syntax As ParserSyntax)
            Me.Syntax = If(syntax, New ParserSyntax())
            AddCurrentResultPair()
        End Sub

        Public ReadOnly Property Syntax As ParserSyntax

        ''' <summary>Problems found while loading (bad entries, reserved prefixes).</summary>
        Public ReadOnly Property Warnings As IReadOnlyList(Of String)
            Get
                Return _warnings
            End Get
        End Property

        ''' <summary>
        ''' False when the map defines the field token (<c>*</c>) as a prefix, which turns off
        ''' deriving field names from the text for empty path fields.
        ''' </summary>
        Public ReadOnly Property WildCardSupportEnabled As Boolean = True

        ''' <summary>The map text this map was loaded from.</summary>
        Public ReadOnly Property SourceText As String
            Get
                Return _sourceText
            End Get
        End Property

        Public ReadOnly Property Definitions As IReadOnlyCollection(Of TokenDefinition)
            Get
                Return _definitions.Values
            End Get
        End Property

#Region "Loading"

        ''' <summary>Builds a map from map text.</summary>
        Public Shared Function Parse(mapText As String, Optional syntax As ParserSyntax = Nothing) As TokenMap
            Dim map = New TokenMap(syntax)
            map.Load(mapText)
            Return map
        End Function

        ''' <summary>Builds a map from a map file.</summary>
        Public Shared Function FromFile(path As String, Optional syntax As ParserSyntax = Nothing) As TokenMap
            Return Parse(File.ReadAllText(path), syntax)
        End Function

        ''' <summary>Builds one of the maps in <see cref="BuiltInMaps"/>.</summary>
        Public Shared Function BuiltIn(name As String, Optional syntax As ParserSyntax = Nothing) As TokenMap
            Return Parse(BuiltInMaps.GetMapText(name), syntax)
        End Function

        ''' <summary>Adds the entries in <paramref name="mapText"/>, replacing existing prefixes.</summary>
        Public Sub Load(mapText As String)
            _sourceText &= mapText
            Dim lines = (If(mapText, "")).Replace(vbCr, "").Split(ChrW(10)).
                Where(Function(l) l.Length > 0 AndAlso Not l.StartsWith(CommentMarker, StringComparison.Ordinal)).ToList()

            Dim i = 0
            While i < lines.Count
                Dim prefix = PrefixOf(lines(i))
                If prefix Is Nothing Then
                    _warnings.Add($"Bad token map line (no '='): [{lines(i)}]")
                    i += 1
                    Continue While
                End If

                Dim openToken = ValueOf(lines(i))
                Dim closeToken = ""
                Dim io = ""
                If i + 1 < lines.Count AndAlso PrefixOf(lines(i + 1)) = prefix Then
                    closeToken = ValueOf(lines(i + 1))
                    If i + 2 < lines.Count AndAlso PrefixOf(lines(i + 2)) = prefix Then
                        io = ValueOf(lines(i + 2))
                        i += 3
                    Else
                        i += 2
                    End If
                Else
                    i += 1
                End If

                If openToken.Length = 0 OrElse closeToken.Length = 0 Then
                    _warnings.Add($"Bad token map entry: [{prefix}:{openToken}:{closeToken}:{io}]")
                    Continue While
                End If

                Add(New TokenDefinition(prefix, openToken, closeToken, io))
            End While
            _prefixes = Nothing
        End Sub

        ''' <summary>Adds or replaces one entry. Returns False for the reserved current-result prefix.</summary>
        Public Function Add(definition As TokenDefinition) As Boolean
            If definition.Prefix = Syntax.CurrentResultToken Then
                _warnings.Add($"Reserved prefix '{definition.Prefix}' cannot be redefined.")
                Return False
            End If
            If definition.Prefix = Syntax.FieldToken Then
                _warnings.Add($"Prefix '{definition.Prefix}' is the field token: wildcard field names are disabled for this map.")
                _WildCardSupportEnabled = False
            End If
            _definitions(definition.Prefix) = definition
            _prefixes = Nothing
            Return True
        End Function

        Private Sub AddCurrentResultPair()
            Dim token = Syntax.CurrentResultToken
            _definitions(token) = New TokenDefinition(token, token, token, "O")
        End Sub

        Private Shared Function PrefixOf(line As String) As String
            Dim eq = line.IndexOf("="c)
            Return If(eq < 0, Nothing, line.Substring(0, eq))
        End Function

        ''' <summary>Map lines starting with this are comments.</summary>
        Public Const CommentMarker As String = "##"

        ''' <summary>The token after the first <c>=</c>, with <c>\n \r \t \\</c> escapes decoded.</summary>
        Private Shared Function ValueOf(line As String) As String
            Return Unescape(line.Substring(line.IndexOf("="c) + 1))
        End Function

        Friend Shared Function Unescape(value As String) As String
            If value.IndexOf("\"c) < 0 Then Return value
            Dim sb = New StringBuilder(value.Length)
            Dim i = 0
            While i < value.Length
                Dim c = value(i)
                If c = "\"c AndAlso i + 1 < value.Length Then
                    Select Case value(i + 1)
                        Case "n"c : sb.Append(vbLf) : i += 2 : Continue While
                        Case "r"c : sb.Append(vbCr) : i += 2 : Continue While
                        Case "t"c : sb.Append(vbTab) : i += 2 : Continue While
                        Case "s"c : sb.Append(" "c) : i += 2 : Continue While
                        Case "\"c : sb.Append("\"c) : i += 2 : Continue While
                    End Select
                    Dim literal = LiteralChars.IndexOf(value(i + 1))
                    If literal >= 0 Then
                        ' an escaped special character: stored as a placeholder the token parser ignores
                        sb.Append(ChrW(LiteralBase + literal))
                        i += 2
                        Continue While
                    End If
                End If
                sb.Append(c)
                i += 1
            End While
            Return sb.ToString()
        End Function

        ''' <summary>Characters that can be escaped with <c>\</c> to be matched literally in a token.</summary>
        Friend Const LiteralChars As String = "*^|?%"
        Friend Const LiteralBase As Integer = &HE000

        ''' <summary>Turns escaped-literal placeholders back into the characters they stand for.</summary>
        Friend Shared Function Literal(value As String) As String
            Dim needs = False
            For Each c In value
                If AscW(c) >= LiteralBase AndAlso AscW(c) < LiteralBase + LiteralChars.Length Then
                    needs = True
                    Exit For
                End If
            Next
            If Not needs Then Return value
            Dim chars = value.ToCharArray()
            For i = 0 To chars.Length - 1
                Dim code = AscW(chars(i)) - LiteralBase
                If code >= 0 AndAlso code < LiteralChars.Length Then chars(i) = LiteralChars(code)
            Next
            Return New String(chars)
        End Function

        Friend Shared Function Escape(value As String) As String
            Dim s = value.Replace("\", "\\").Replace(vbLf, "\n").Replace(vbCr, "\r").Replace(vbTab, "\t")
            ' keep leading/trailing spaces visible so editors cannot strip them
            If s.EndsWith(" "c) Then s = s.Substring(0, s.Length - 1) & "\s"
            If s.StartsWith(" "c) Then s = "\s" & s.Substring(1)
            For i = 0 To LiteralChars.Length - 1
                s = s.Replace(ChrW(LiteralBase + i).ToString(), "\" & LiteralChars(i))
            Next
            Return s
        End Function

#End Region

#Region "Lookup"

        Private Shared Function Normalize(prefix As String) As String
            Return If(String.IsNullOrEmpty(prefix), ParserSyntax.DefaultPrefix, prefix)
        End Function

        Public Function Contains(prefix As String) As Boolean
            Return _definitions.ContainsKey(Normalize(prefix))
        End Function

        ''' <summary>A fresh token pair for <paramref name="prefix"/>, or Nothing if it is not mapped.</summary>
        Public Function CreatePair(prefix As String) As TokenPair
            Dim definition As TokenDefinition = Nothing
            If _definitions.TryGetValue(Normalize(prefix), definition) Then
                Return New TokenPair(Syntax, definition)
            End If
            Return Nothing
        End Function

        ''' <summary>All prefixes, longest first, so path segments match the most specific prefix.</summary>
        Public ReadOnly Property Prefixes As IReadOnlyList(Of String)
            Get
                If _prefixes Is Nothing Then
                    _prefixes = _definitions.Keys.
                        OrderByDescending(Function(p) p.Length).
                        ThenBy(Function(p) p, StringComparer.Ordinal).
                        ToArray()
                End If
                Return _prefixes
            End Get
        End Property

        ''' <summary>The map as map text, one line per token.</summary>
        Public Overrides Function ToString() As String
            Dim sb = New StringBuilder()
            For Each d In _definitions.Values
                If d.Prefix = Syntax.CurrentResultToken Then Continue For
                sb.Append(d.Prefix).Append("="c).AppendLine(Escape(d.OpenToken))
                sb.Append(d.Prefix).Append("="c).AppendLine(Escape(d.CloseToken))
                sb.Append(d.Prefix).Append("="c).AppendLine(d.IOSpec)
            Next
            Return sb.ToString()
        End Function

#End Region

    End Class

End Namespace
