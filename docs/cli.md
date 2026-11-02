# Command line

```text
objParser [options] -x <path> [-x <path> ...]
objParser <path>                          (one bare argument: a path, input from stdin)
objParser [options] -d <open> -D <close>  (ad-hoc delimiters instead of a path)
```

Get a free-standing `objParser.exe` (one file, runs without .NET installed) with
`dotnet publish src/objParser -p:PublishProfile=win-x64`; it is written to
`artifacts/win-x64/objParser.exe`.

Results go to **stdout**, one per line. Diagnostics, timings, counts and warnings go to **stderr**,
so they never mix with results in a pipe.

## Options

### Input and maps

| Option | Meaning |
|---|---|
| `-f <file>` | read this file; without it, read stdin |
| `-xml` `-html` `-xhtml` | the XML map (the default) |
| `-xml_ext` | XML plus named prefixes |
| `-vb` | VB.NET map |
| `-cs` `-csharp` | C# map |
| `-m <mapfile>` | use a [token map file](token-maps.md) instead of a built-in map |
| `-showmap` | print the selected map and exit |

### Queries

| Option | Meaning |
|---|---|
| `-x <path>` | a [path](path-syntax.md) to evaluate; repeat for more, and results come in the order given |
| `-d <open> -D <close>` | find text between two tokens, without a map; `~` stands for `"` |
| `-s` | sort the results and drop duplicates |
| `-results_token <t>` | use `<t>` instead of `.` as the current-result step |

### Editing

With these options the tool prints the **edited text** instead of results.

| Option | Meaning |
|---|---|
| `-delete` | remove every match of the `-x` paths (applied one path after another) |
| `-replace <text>` | with `-delete`: put `<text>` in place of each match |
| `-keep <path>` | with `-delete`: put the results of `<path>`, evaluated inside the match, in its place |
| `-clean <path>` | remove matches before any query runs (repeatable); good for stripping comments |

### Diagnostics

| Option | Meaning |
|---|---|
| `-count` | report the number of results |
| `-t` | report the time per path segment, for parsing and for filtering |
| `-v` | report paths and filters as they run, plus map warnings |
| `-debug` | report every token search and where it matched |
| `-ver` | print the version |
| `-help` | print usage |

## Exit codes

| Code | Meaning |
|---|---|
| 0 | success (finding no results is still a success) |
| 1 | the input file or map file was not found |
| 2 | the command line was wrong: unknown option, missing value, or a missing `-x` |

## Examples

```console
# attribute values, sorted and unique
objParser -f model.xml -x "//flags/flag/@attr" -s

# two queries, results concatenated
objParser -f model.xml -x "//decision/@name" -x "//variable/@name"

# from a pipe
type page.html | objParser -html "//a/@href"

# strip every <index> element and print the rest of the document
objParser -f model.xml -x "//index" -delete > trimmed.xml

# replace each variable with just its name
objParser -f model.xml -x "//variable" -delete -keep "@name"

# ignore anything inside comments, then query
objParser -xml_ext -f model.xml -clean "//comment:" -x "//@name"

# a quick one-off without a map
objParser -f model.xml -d "name=~" -D "~"

# your own map
objParser -m examples/mediawiki.map -f examples/release.wiki -x "//Owners/li"
```

### Running from Git Bash

Git Bash rewrites arguments that look like Unix paths (for example `-D "</node>"`). Turn that off
with `MSYS_NO_PATHCONV=1 objParser ...`. PowerShell and cmd don't do this.
