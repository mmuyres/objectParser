Namespace ObjectParser

    ''' <summary>
    ''' The token maps that ship with the parser, in map-file format (run <c>objParser -showmap -xml</c>
    ''' to print one). Each is a working example of how to describe a format.
    ''' </summary>
    Public NotInheritable Class BuiltInMaps

        Private Sub New()
        End Sub

        ''' <summary>XML, XHTML and HTML.</summary>
        Public Const Xml As String =
"## XML / XHTML / HTML
## name     whole element, tags included        //variable
 =<*^>
 =/>?/>|</*^>
 =O|O
## -name    element content, tags excluded      //-title
-=<*^>
-=/>?/>|</*^>
-=I|I
## @name    attribute value; @ alone = all      //variable/@name
@=*=""
@=""
@=I
## $        attribute names                     //variable/$
$=*=""
$=""
$=F
## #        element names                       //variable/#
#=<*
#=|>|/
#=F
## !        first text run between > and <      //choice/!
!=>
!=<
!=I
"

        ''' <summary>The XML map plus named prefixes for comments, CDATA, processing instructions and tags.</summary>
        Public Const XmlExtended As String = Xml &
"## inner:    same as !
inner:=>
inner:=<
inner:=I
## node:name   opening tag only                 //node:variable
node:=<*
node:=/>?/>|>
node:=O
## attname:    attribute names (same as $)
attname:=*=""
attname:=""
attname:=F
## nodename:   element names (same as #)
nodename:=<*
nodename:=|>|/
nodename:=F
## comment:    <!-- comments --> (whole comment)
comment:=<!--
comment:=-->
comment:=O
## cdata:      CDATA section content
cdata:=<![CDATA[
cdata:=]]>
cdata:=I
## pi:         processing instruction content   //pi:
pi:=<\?
pi:=\?>
pi:=I
"

        ''' <summary>VB.NET source.</summary>
        Public Const VisualBasic As String =
"## VB.NET source
## Name     class block                         //Invoice
 =Class *
 =End Class
 =O
## -Name    class body                          //-Invoice
-=Class *
-=End Class
-=I
## !Name    function block; ! alone = all       //Invoice/!Total
!=Function *^(
!=End Function
!=O
## @Name    sub block; @ alone = all            //Invoice/@Reset
@=Sub *^(
@=End Sub
@=O
## :        class names                         //:
:=Class *
:=|\n| |(|\r
:=F
## p        function names                      //Invoice/p
p=Function *
p=(
p=F
## $        sub names                           //Invoice/$
$=Sub *
$=(
$=F
## f        function signatures                 //f
f=Function ^(
f=)
f=O
## s        sub signatures                      //s
s=Sub ^(
s=)
s=O
## ;        ByVal parameter names               //Invoice/@AddLine/;
;=ByVal\s
;= As
;=I
## v        names declared with As              //Invoice/v
v=* As\s
v=|\n| |,|)|\r
v=F
## #        type names after As                 //#
#= As *
#=|\n| |,|)|(|\r
#=F
## m        member calls: Name in obj.Name(     //m
m=*.
m=(
m=I
"

        ''' <summary>C# source. Block prefixes use balanced braces (<c>%{};</c>).</summary>
        Public Const CSharp As String =
"## C# source
## Name     class block                         //Invoice
 =class *
 =%{};
 =O
## -Name    class body, inside the braces       //-Invoice
-=class *
-=%{};
-=I
## iName    interface block                     //iIShape
i=interface *
i=%{};
i=O
## sName    struct block                        //sPoint
s=struct *
s=%{};
s=O
## eName    enum body                           //eColor
e=enum *
e=%{};
e=I
## !Name    method or constructor, name to closing brace; ! alone = all
!=*(
!=%{};
!=O
## m        method and constructor names        //Invoice/m
m=*(
m=%{};
m=F
## (        parenthesised groups in the block; in a method block the first is its parameter list
(=(
(=)
(=I
## :        class names                         //:
:=class *
:=|{|:|<|;| |\n
:=F
## n        namespace names                     //n
n=namespace *
n=|{|;|\n|\r
n=F
## u        using directives                    //u
u=using *
u=;
u=F
## c        // line comment text                //c
c=//
c=|\r|\n
c=I
## b        /* block comment */ text            //b
b=/\*
b=\*/
b=I
## #        #region names                       //#
#=#region\s
#=|\r|\n
#=I
"

        ''' <summary>The names accepted by <see cref="GetMapText"/>.</summary>
        Public Shared ReadOnly Property Names As IReadOnlyList(Of String) = {"xml", "xml_ext", "html", "xhtml", "vb", "cs", "csharp"}

        ''' <summary>Map text for a built-in map (case-insensitive). <c>html</c>/<c>xhtml</c> use the XML map; <c>cs</c> = <c>csharp</c>.</summary>
        Public Shared Function GetMapText(name As String) As String
            Select Case If(name, "").ToLowerInvariant()
                Case "xml", "html", "xhtml"
                    Return Xml
                Case "xml_ext"
                    Return XmlExtended
                Case "vb"
                    Return VisualBasic
                Case "cs", "csharp", "c#"
                    Return CSharp
                Case Else
                    Throw New ArgumentException($"Unknown built-in map '{name}'. Known maps: {String.Join(", ", Names)}.", NameOf(name))
            End Select
        End Function

    End Class

End Namespace
