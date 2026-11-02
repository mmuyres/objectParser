# Built-in maps

Choose one with a CLI switch or `TokenMap.BuiltIn("<name>")`. Print its map text with
`objParser -showmap -<name>`. The source is
[`src/ObjectParser/BuiltInMaps.vb`](../src/ObjectParser/BuiltInMaps.vb), written in map-file
format, so each map doubles as an example.

| Switch | Map | For |
|---|---|---|
| `-xml` (default), `-html`, `-xhtml` | `xml` | XML and HTML |
| `-xml_ext` | `xml_ext` | XML plus named prefixes for comments, CDATA, tags |
| `-vb` | `vb` | VB.NET source |
| `-cs`, `-csharp` | `cs` | C# source |

In the tables, *Name* stands for the name you put in the path.

## xml

| Path | Returns | Scope |
|---|---|---|
| `//Name` | the whole element, tags included | O |
| `//-Name` | the element's content | I |
| `//@Name` | an attribute value; `@` alone gives every attribute value | I |
| `//$` | attribute names | F |
| `//#` | element names, from opening tags (self-closing included) | F |
| `//!` | the first run of text between `>` and `<` | I |
| `//Name/` | the element's direct children (an empty name is a wildcard) | O |

## xml_ext

Everything in `xml`, plus:

| Path | Returns |
|---|---|
| `//comment:` | `<!-- comments -->`, whole |
| `//cdata:` | CDATA section content |
| `//pi:` | processing instruction content (`xml version="1.0"`) |
| `//node:Name` | the opening tag only |
| `//nodename:` | element names (same as `#`) |
| `//attname:` | attribute names (same as `$`) |
| `//inner:` | same as `!` |

## vb

| Path | Returns |
|---|---|
| `//Name` | a class, `Class Name` … `End Class` |
| `//-Name` | the class body |
| `//:` | class names |
| `//Class/!Name` | a function, `Function Name(` … `End Function`; `!` alone gives every function |
| `//Class/@Name` | a sub, `Sub Name(` … `End Sub`; `@` alone gives every sub |
| `//p` | function names |
| `//$` | sub names |
| `//f`, `//s` | function and sub signatures, up to the closing `)` |
| `//;` | `ByVal` parameter names |
| `//v` | names declared with `As` (fields, parameters, locals) |
| `//#` | type names after `As` |
| `//m` | member calls: `Trim` in `first.Trim(` |

```console
$ objParser -vb -f Sample.vb -x "//Invoice/v"
m_total
amount
qty
```

## cs

C# blocks are found with balanced braces, so nested braces, `if` blocks and lambdas don't cut a
class or method short. The `;` stop character keeps calls such as `Add(x);` from being taken as
declarations.

| Path | Returns |
|---|---|
| `//Name` | a class, `class Name` through its closing brace |
| `//-Name` | the class body, inside the braces |
| `//:` | class names |
| `//iName`, `//sName` | an interface or struct block |
| `//eName` | an enum body |
| `//Class/!Name` | a method or constructor, from its name through its closing brace; `!` alone gives every one |
| `//Class/m` | method and constructor names |
| `//(` | parenthesised groups; in a method block, the first is the parameter list |
| `//n` | namespace names |
| `//u` | `using` directives |
| `//c` | `//` comment text |
| `//b` | `/* */` comment text |
| `//#` | `#region` names |

```console
$ objParser -cs -f Sample.cs -x "//Invoice/m"
Invoice
Add
Total
Dispose
```

Limits: braces inside strings or comments are counted like any other, expression-bodied members
(`=> x;`) are not found by `!`, and a keyword written without a space before its parenthesis
(`if(x) {`) can show up as a method name.

## Known limits of the XML map

- Elements nested inside an element of the same name are not supported.
- `//Name/-Name` finds nothing; see [path syntax](path-syntax.md#searches-go-through-the-whole-block).
- `@Name` and `$` also match attribute-like text inside comments or CDATA. Strip those first with
  `-clean "//comment:"` (using `-xml_ext`).
