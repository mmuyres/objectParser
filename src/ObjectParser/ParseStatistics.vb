Imports System.Globalization
Imports System.Text

Namespace ObjectParser

    ''' <summary>Timing collected while a <see cref="Parser"/> evaluates paths.</summary>
    Public NotInheritable Class ParseStatistics

        Private ReadOnly _segments As New List(Of KeyValuePair(Of String, Double))

        ''' <summary>Time per path segment, in evaluation order.</summary>
        Public ReadOnly Property Segments As IReadOnlyList(Of KeyValuePair(Of String, Double))
            Get
                Return _segments
            End Get
        End Property

        Public Property ParseMilliseconds As Double
        Public Property FilterMilliseconds As Double

        Friend Sub AddSegment(segment As String, milliseconds As Double)
            _segments.Add(New KeyValuePair(Of String, Double)(segment, milliseconds))
        End Sub

        ''' <summary>Folds a nested parser's time into this one (filters, sub-paths).</summary>
        Friend Sub Add(other As ParseStatistics, includeSegments As Boolean)
            ParseMilliseconds += other.ParseMilliseconds
            FilterMilliseconds += other.FilterMilliseconds
            If includeSegments Then _segments.AddRange(other._segments)
        End Sub

        Public Overrides Function ToString() As String
            Dim inv = CultureInfo.InvariantCulture
            Dim width = Math.Max(12, If(_segments.Count = 0, 0, _segments.Max(Function(s) s.Key.Length)))
            Dim sb = New StringBuilder()
            sb.AppendLine("Path Segment".PadLeft(width) & "      Time")
            For Each s In _segments
                sb.AppendLine(s.Key.PadLeft(width) & "  " & s.Value.ToString("0.000", inv).PadLeft(8) & " ms")
            Next
            sb.AppendLine("Parsing".PadLeft(width) & "  " & ParseMilliseconds.ToString("0.000", inv).PadLeft(8) & " ms")
            sb.AppendLine("Filtering".PadLeft(width) & "  " & FilterMilliseconds.ToString("0.000", inv).PadLeft(8) & " ms")
            Return sb.ToString()
        End Function

    End Class

End Namespace
