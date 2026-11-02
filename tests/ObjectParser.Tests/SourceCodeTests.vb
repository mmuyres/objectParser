Imports ObjectParser
Imports Xunit

''' <summary>The vb and cs built-in maps against the sample source files.</summary>
Public Class SourceCodeTests

    Private Shared ReadOnly Vb As String = ReadData("Sample.vb")
    Private Shared ReadOnly Cs As String = ReadData("Sample.cs")

    Private Shared Function V(path As String) As List(Of String)
        Return LfAll(Q(Vb, path, "vb"))
    End Function
    Private Shared Function C(path As String) As List(Of String)
        Return LfAll(Q(Cs, path, "cs"))
    End Function

#Region "VB"

    <Fact>
    Public Sub Vb_class_names()
        Assert.Equal({"Invoice", "Customer"}, V("//:"))
    End Sub

    <Fact>
    Public Sub Vb_class_block_by_name()
        Dim invoice = V("//Invoice").Single()
        Assert.StartsWith("Class Invoice", invoice)
        Assert.EndsWith("End Class", invoice)
        Assert.DoesNotContain("Customer", invoice)
    End Sub

    <Fact>
    Public Sub Vb_member_names_within_a_class()
        Assert.Equal({"Total"}, V("//Invoice/p"))
        Assert.Equal({"AddLine", "Reset"}, V("//Invoice/$"))
        Assert.Equal({"DisplayName"}, V("//Customer/p"))
    End Sub

    <Fact>
    Public Sub Vb_function_and_sub_blocks()
        Assert.Equal({"Function Total() As Decimal" & vbLf & "        Return m_total" & vbLf & "    End Function"}, V("//Invoice/!Total"))
        Assert.Equal(2, V("//Invoice/@").Count)
        Assert.All(V("//Invoice/@"), Sub(s) Assert.EndsWith("End Sub", s))
        Assert.Equal(2, V("//!").Count) ' every function in the file
    End Sub

    <Fact>
    Public Sub Vb_signatures_parameters_and_types()
        Assert.Equal({"Sub AddLine(ByVal amount As Decimal, ByVal qty As Integer)", "Sub Reset()"}, V("//Invoice/s"))
        Assert.Equal({"amount", "qty"}, V("//Invoice/@AddLine/;"))
        Assert.Equal({"m_total", "amount", "qty"}, V("//Invoice/v"))
        Assert.Equal({"Decimal", "Decimal", "Integer", "Decimal", "String", "String"}, V("//#"))
        Assert.Equal({"Trim"}, V("//m"))
    End Sub

#End Region

#Region "C#"

    <Fact>
    Public Sub Cs_class_names_and_blocks()
        Assert.Equal({"Invoice", "InvoiceLine"}, C("//:"))
        Dim invoice = C("//Invoice").Single()
        Assert.StartsWith("class Invoice : IDisposable", invoice)
        Assert.EndsWith("public void Dispose() { _lines.Clear(); }" & vbLf & "    }", invoice)
        Assert.Equal({"public decimal Amount { get; set; }"}, C("//-InvoiceLine").Select(Function(s) s.Trim()))
    End Sub

    <Fact>
    Public Sub Cs_method_names_skip_calls_and_keywords()
        Assert.Equal({"Invoice", "Add", "Total", "Dispose"}, C("//Invoice/m"))
    End Sub

    <Fact>
    Public Sub Cs_method_block_and_parameters()
        Dim add = C("//Invoice/!Add").Single()
        Assert.StartsWith("Add(decimal amount, int qty)", add)
        Assert.EndsWith("}", add)
        Assert.Equal("decimal amount, int qty", C("//Invoice/!Add/(").First())
    End Sub

    <Fact>
    Public Sub Cs_other_type_blocks()
        Assert.StartsWith("interface IShape", C("//iIShape").Single())
        Assert.Equal({"Red, Green, Blue"}, C("//eColor").Select(Function(s) s.Trim()))
        Assert.Contains("public int X;", C("//sPoint").Single())
    End Sub

    <Fact>
    Public Sub Cs_namespaces_usings_comments_regions()
        Assert.Equal({"Billing.Core"}, C("//n"))
        Assert.Equal({"System", "System.Collections.Generic"}, C("//u"))
        Assert.Equal({"An invoice with line items."}, C("//c").Select(Function(s) s.Trim()))
        Assert.Equal({"Billing types used by the samples."}, C("//b").Select(Function(s) s.Trim()))
        Assert.Equal({"Totals"}, C("//#"))
    End Sub

    <Fact>
    Public Sub Cs_generic_and_partial_names_use_word_boundaries()
        Dim code = "class Box<T> { } class BoxSet { } partial class Box2 { }"
        Assert.Equal({"Box", "BoxSet", "Box2"}, Q(code, "//:", "cs"))
        Assert.Equal({"class Box<T> { }"}, Q(code, "//Box", "cs"))
    End Sub

#End Region

End Class
