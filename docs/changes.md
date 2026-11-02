# Changes since v4.1

v4.1 (2010) was a single Visual Studio 2008 console project. v5 (2026) is a .NET 9 solution with
a reusable library, a command-line tool and a test suite. Before porting, the original executable
was rebuilt and run against the same sample files, and every difference below is deliberate.

The v4.1 sources are kept under the git tag `v4.1-legacy`.

## Structure

| v4.1 | v5 |
|---|---|
| one `.vbproj`, `Option Strict Off`, `Hashtable`/`ArrayList` | `src/ObjectParser` library + `src/objParser` CLI, `Option Strict On`, generic collections |
| the library printed to the console | the library returns results; logging goes to a `TextWriter` you supply |
| no tests | 517 xUnit tests (see below) |
| built binaries committed in `bin/` | built by `dotnet build`, ignored by git |

## Fixed

| Problem in v4.1 | Example | Now |
|---|---|---|
| results came back in hash-table order | `//variable` gave alpha, gamma, beta | document order |
| element names matched longer names | `//choice` also returned `<choices>` | names stop at a word boundary (also `@id` vs `xid`) |
| `[!x]` (documented) did an existence check | `//variable[!flags]` gave the variables *with* flags | means "x is absent" |
| a right-hand path with a single value matched nothing | `//choices/choice[@type=//controls/type[@value='pulldown']/@value]` returned no results | works |
| `!` and `~` against a list were inverted | `[@type!//controls/type/@value]` | follow the documented meaning |
| `-delete` with `-x` printed an empty document | `-x //flags -delete -replace X` printed only `X` | deletes each match; several `-x` apply in turn |
| deletes used `String.Replace` on the matched text | deleting attribute value `1` removed every `1` | deletes cut out exact positions |
| an empty element stopped the scan | `<td></td><td>x</td>` lost `x` | continues |
| filters and sub-paths reused the outer query's token objects | wrong offsets after a filter, wildcard state that never cleared | fresh token pairs per use |
| wildcard names only worked for one-character leads | `//Invoice/!` (vb) found nothing | general rule, works for any lead |
| a missing input file or option value crashed | `-f nofile.xml`, a trailing `-x` | a message and an exit code |
| a map's `/>` look-ahead length was reset to 2 on every lookup | | kept as configured |
| a search-conditional close token chose a "not found" alternative | `-1` won the "nearest" comparison | only found alternatives count |

## Changed

- `-html` now uses the XML map. The old HTML map was a reduced copy without `-` inner text.
- The XML map no longer defines `*` (closing-tag names). That prefix switched off wildcards for
  the whole map. Element names are now `#`, attribute names `$` (both with the new `F` scope), and
  `//Name/` returns children.
- VB map: `@` (subs) returns the whole sub like `!` does for functions. `:`, `p`, `$`, `v`, `#`
  return clean names. `p` no longer runs from `End Function` into the next method.
- Default output no longer silently drops a result that equals the one before it. Use `-s` for
  unique output.
- A leading `-` in a filter is now the inner-text prefix (v4.1 stripped it as an unimplemented
  "inner filter" marker). `[+x]` is still accepted.
- Warnings for the built-in maps are no longer printed on every run. Warnings for `-m` map files
  go to stderr.

## Added

- **C# map** (`-cs`): classes, interfaces, structs, enums, methods, parameters, namespaces, usings,
  comments, regions.
- **`F` scope**: return the name `*` matched (element, attribute, class, method names).
- **Balanced tokens** `%{}`, `%()`, …, with an optional stop character (`%{};`).
- **Lookahead close tokens**: `|a|b` and `^a` with nothing in front end a match without consuming
  the target, and also match at end of text.
- Map files: `##` comments and escapes `\n \r \t \s \\ \* \^ \| \? \%`.
- `xml_ext`: `cdata:` and `pi:` prefixes.
- Library API: `Query`, `QueryMatches` (with positions), `QueryAll`, `Delete`, `Clean`,
  `FindBetween`, `TokenMap.Parse/FromFile/BuiltIn`.
- CLI: `-cs`, `-help`, exit codes (0 / 1 input / 2 usage).

## How it was tested

| Suite | What it proves |
|---|---|
| `XmlQueryTests`, `FilterTests` | every example from the v4.1 usage text, against `examples/model.xml` |
| `CrossCheckTests` | 300 random well-formed documents (fixed seeds): elements, inner text, attributes, filters, deletes and names all agree with what the generator wrote and with `System.Xml.Linq`, down to the character position |
| `RobustnessTests` | random and truncated input with every built-in map never hangs, never throws, and every reported match is really at its position; 20,000 elements in well under a second |
| `TokenGrammarTests`, `TokenMapTests`, `PathExpressionTests` | each token form, escape and path feature in isolation |
| `SourceCodeTests`, `WikiAndHtmlTests`, `CustomMapTests` | the vb and cs maps, wiki pages as HTML and MediaWiki, plus Markdown, INI, log and JSON-like maps |
| `EditingTests`, `CliTests` | delete/replace/keep/clean, and the CLI's output, errors and exit codes |
