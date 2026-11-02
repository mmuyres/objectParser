Imports System.Text
Imports System.Xml.Linq
Imports ObjectParser
Imports Xunit

''' <summary>
''' Proof by comparison: generate random well-formed XML (fixed seeds, so runs are repeatable),
''' record where every element and attribute was written, confirm the document with
''' System.Xml.Linq, then check that objParser finds exactly the same things at exactly the same
''' offsets. Names deliberately share prefixes (item/items, id/xid) to exercise word boundaries.
''' </summary>
Public Class CrossCheckTests

    Private Shared ReadOnly ElementNames As String() = {"a", "b", "item", "items", "data", "d"}
    Private Shared ReadOnly AttributeNames As String() = {"id", "xid", "name", "k"}

    Private NotInheritable Class GenElement
        Public Name As String
        Public OuterStart As Integer
        Public OuterEnd As Integer
        Public InnerStart As Integer
        Public InnerEnd As Integer
        Public SelfClosing As Boolean
        Public ReadOnly Attributes As New List(Of GenAttribute)
    End Class

    Private NotInheritable Class GenAttribute
        Public Name As String
        Public Value As String
        Public ValueStart As Integer
    End Class

    Private NotInheritable Class GenDoc
        Public Text As String
        Public ReadOnly Elements As New List(Of GenElement)
        Public ReadOnly Attributes As New List(Of GenAttribute)
    End Class

    Private Shared Function Generate(seed As Integer) As GenDoc
        Dim rnd = New Random(seed)
        Dim doc = New GenDoc()
        Dim sb = New StringBuilder()
        Dim word = Function() New String(Enumerable.Range(0, rnd.Next(1, 6)).Select(Function(i) "abcxyz019"(rnd.Next(9))).ToArray())

        Dim writeElement As Action(Of Integer, HashSet(Of String)) = Nothing
        writeElement = Sub(depth As Integer, ancestors As HashSet(Of String))
                           Dim choices = ElementNames.Where(Function(n) Not ancestors.Contains(n)).ToArray()
                           Dim el = New GenElement With {.Name = choices(rnd.Next(choices.Length)), .OuterStart = sb.Length}
                           doc.Elements.Add(el)
                           sb.Append("<"c).Append(el.Name)
                           For Each attrName In AttributeNames.OrderBy(Function(x) rnd.Next()).Take(rnd.Next(0, 3))
                               sb.Append(If(rnd.Next(3) = 0, vbLf & "  ", " ")).Append(attrName).Append("=""")
                               Dim attr = New GenAttribute With {.Name = attrName, .Value = word(), .ValueStart = sb.Length}
                               sb.Append(attr.Value).Append(""""c)
                               el.Attributes.Add(attr)
                               doc.Attributes.Add(attr)
                           Next
                           Dim childCount = If(depth >= 4, 0, rnd.Next(0, 4))
                           If childCount = 0 AndAlso rnd.Next(2) = 0 Then
                               el.InnerStart = sb.Length
                               sb.Append("/>")
                               el.SelfClosing = True
                               el.InnerEnd = sb.Length - 2
                           Else
                               sb.Append(">"c)
                               el.InnerStart = sb.Length
                               Dim inner = New HashSet(Of String)(ancestors) From {el.Name}
                               For c = 1 To childCount
                                   If rnd.Next(2) = 0 Then sb.Append(word())
                                   If rnd.Next(3) = 0 Then sb.Append(vbLf)
                                   writeElement(depth + 1, inner)
                               Next
                               If childCount = 0 OrElse rnd.Next(2) = 0 Then sb.Append(word())
                               el.InnerEnd = sb.Length
                               sb.Append("</").Append(el.Name).Append(">"c)
                           End If
                           el.OuterEnd = sb.Length
                       End Sub

        sb.Append("<root>")
        For i = 1 To rnd.Next(1, 6)
            writeElement(1, New HashSet(Of String)({"root"}))
        Next
        sb.Append("</root>")
        doc.Text = sb.ToString()
        Return doc
    End Function

    Public Shared ReadOnly Property Seeds As IEnumerable(Of Object())
        Get
            Return Enumerable.Range(1, 300).Select(Function(s) New Object() {s})
        End Get
    End Property

    <Theory>
    <MemberData(NameOf(Seeds))>
    Public Sub Generated_documents_agree_with_the_generator(seed As Integer)
        Dim doc = Generate(seed)

        ' the generator really produced well-formed XML with the elements it recorded
        Dim xdoc = XDocument.Parse(doc.Text, LoadOptions.PreserveWhitespace)
        Assert.Equal(doc.Elements.Count, xdoc.Root.Descendants().Count())

        Dim parser = New Parser(TokenMap.BuiltIn("xml"), doc.Text)
        For Each name In ElementNames
            Dim expected = doc.Elements.Where(Function(e) e.Name = name).ToList()

            ' //name -> every element with that name, outer text, in document order
            Dim outer = parser.QueryMatches("//" & name)
            Assert.Equal(expected.Select(Function(e) e.OuterStart), outer.Select(Function(m) m.Start))
            Assert.Equal(expected.Select(Function(e) doc.Text.Substring(e.OuterStart, e.OuterEnd - e.OuterStart)), outer.Select(Function(m) m.Value))
            Assert.Equal(xdoc.Descendants(name).Count(), outer.Count)

            ' //-name -> content of paired elements (self-closing ones give their attribute text)
            Dim inner = parser.QueryMatches("//-" & name)
            Assert.Equal(expected.Select(Function(e) If(e.SelfClosing, e.OuterStart + 1 + e.Name.Length, e.InnerStart)), inner.Select(Function(m) m.Start))
            For Each pair In expected.Zip(inner)
                If Not pair.First.SelfClosing Then
                    Assert.Equal(doc.Text.Substring(pair.First.InnerStart, pair.First.InnerEnd - pair.First.InnerStart), pair.Second.Value)
                End If
            Next

            For Each attrName In AttributeNames
                ' //name/@attr -> that attribute anywhere inside each such element, by position
                Dim expectedValues = expected.SelectMany(
                    Function(e) doc.Attributes.Where(Function(a) a.Name = attrName AndAlso a.ValueStart > e.OuterStart AndAlso a.ValueStart < e.OuterEnd).
                                               OrderBy(Function(a) a.ValueStart)).ToList()
                Dim actual = parser.QueryMatches("//" & name & "/@" & attrName)
                Assert.Equal(expectedValues.Select(Function(a) a.ValueStart), actual.Select(Function(m) m.Start))
                Assert.Equal(expectedValues.Select(Function(a) a.Value), actual.Select(Function(m) m.Value))

                ' filters agree with LINQ to XML
                If expectedValues.Count > 0 Then
                    Dim probe = expectedValues(0).Value
                    Dim filtered = parser.Query($"//{name}[@{attrName}='{probe}']").Count
                    Dim linq = xdoc.Descendants(name).Count(Function(x) x.DescendantsAndSelf().Any(Function(d) CStr(d.Attribute(attrName)) IsNot Nothing AndAlso
                                                                                                           String.Equals(CStr(d.Attribute(attrName)), probe, StringComparison.OrdinalIgnoreCase)))
                    Assert.Equal(linq, filtered)
                End If
                Dim withAttr = parser.Query($"//{name}[@{attrName}]").Count
                Dim withoutAttr = parser.Query($"//{name}[@{attrName}!]").Count
                Assert.Equal(expected.Count, withAttr + withoutAttr)
                Assert.Equal(xdoc.Descendants(name).Count(Function(x) x.DescendantsAndSelf().Any(Function(d) d.Attribute(attrName) IsNot Nothing)), withAttr)
            Next

            ' deleting every element of a name gives the text with exactly those spans removed
            Dim editor = New Parser(TokenMap.BuiltIn("xml"), doc.Text)
            Dim expectedText = New StringBuilder()
            Dim cursor = 0
            For Each e In expected.Where(Function(x) x.OuterStart >= 0).OrderBy(Function(x) x.OuterStart)
                If e.OuterStart < cursor Then Continue For
                expectedText.Append(doc.Text, cursor, e.OuterStart - cursor)
                cursor = e.OuterEnd
            Next
            expectedText.Append(doc.Text, cursor, doc.Text.Length - cursor)
            Assert.Equal(expectedText.ToString(), editor.Delete("//" & name))
        Next

        ' element names and attribute names, via the F-scope prefixes
        Assert.Equal(doc.Elements.Select(Function(e) e.Name).Prepend("root"), parser.Query("//#"))
        Assert.Equal(doc.Attributes.Select(Function(a) a.Name), parser.Query("//$"))
    End Sub

End Class
