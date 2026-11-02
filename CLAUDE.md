# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

`objParser` is a VB.NET console tool (circa 2010, v4.1) that pulls text out of structured documents using a pseudo-XPath plus a **token map**. It never builds a DOM. Each path segment's prefix names a pair of open/close delimiter strings, and the parser does string searches between them. The same engine works on XML, HTML, VB source, or any format you can describe with delimiters.

## Build

It is a legacy VS2008-era project (`ToolsVersion="3.5"`, no target framework set, so it builds against .NET 2.0/3.5). There are no tests, no linter and no CI.

```sh
# The project references old_key.ico, which is not in the repo, so override ApplicationIcon or the build fails (BC30136)
/c/Windows/Microsoft.NET/Framework/v3.5/msbuild.exe objParser.vbproj //p:ApplicationIcon=
```

- Output goes to `bin/objParser.exe`, and `bin/` is committed. Rebuilding changes tracked binaries, so build to another folder (`//p:OutputPath=... //p:IntermediateOutputPath=...`) unless you mean to update them.
- A warning about the Windows SDK v6.0A missing can be ignored.
- `OptionStrict` is Off, so late binding and implicit conversions are everywhere.
- `Module1.vb` is an empty stub and is **not** in the `.vbproj` compile list. The real entry point is `Module Module1` in `Main.vb` (`StartupObject` is `objParser.Module1`).
- `DefinedMaps.vb` and `xmlTokenMaps.txt` are leftovers or samples, and nothing at runtime uses them.

## Running (manual smoke test)

```sh
echo '<root name="a"><item id="1"/><item id="2"/></root>' | bin/objParser.exe -x "//item/@id"
bin/objParser.exe -f file.xml -x "root/-name" -s -count -t
bin/objParser.exe                # full usage text, which is the authoritative syntax reference
bin/objParser.exe -showmap -xml  # print a built-in token map
```

- With no `-f`, input comes from stdin. A single bare argument is taken as the xpath.
- Results go to stdout. Timings (`-t`), counts (`-count`) and warnings go to stderr.
- The built-in `xml` map defines the `*` prefix, so every run prints a "Reserved Token Prefix … Wild Card support … disabled" warning on stderr. This is expected.
- Results are kept in `Hashtable`s, so output order is not guaranteed (`-s` gives sorted, unique output).

## Architecture

Everything is in `Namespace ObjectParser` except the CLI module.

- **`Main.vb`**: parses CLI arguments in two passes. The second pass re-reads arguments that can repeat (`-x`, `-clean`) once the `Parser` exists. It also holds the built-in token maps (`makeTokenMap`: `xml`, `xml_ext`, `html`, `vb`) and `ShowUsage`, which documents the token-map and xpath syntax in detail. Read `ShowUsage` before changing either syntax.
- **`ParserTokenType`**: one config object that is created in `Main`, then **cloned** into each `TokenMap`, `Parser` and `xPath`. It holds:
  - bit-flag constants (`JUMP`, `CONDITIONAL`, `INNER_RESULTS`, …) used as a `TypeMask`
  - every special-character setting: `.` current result, `^` jump, `?`/`|` conditional, `~` quote, `{{ }}` sub-path, `{: :}` multi-path, `[ ]` filter, `*` field token, and so on

  Because each object holds a clone, change these settings (for example `-results_token`) **before** you build the objects that clone them.
- **`TokenMap.vb`**:
  - `TokenMap` loads map text as **triplets of lines** per prefix: `prefix=<open token>`, `prefix=<close token>`, `prefix=<I|O>`. `I` or `O` selects inner or outer result text, and conditional tokens can have several `|`-separated values.
  - It builds one `TokenPair` per prefix and always adds a pair for the current-result token `.`.
  - `GetPreFixList()` returns the prefixes sorted in reverse, so `xPath` can match the longest prefix first.
- **`ParserToken.vb`**: one open or close token. It parses the token string into its parts (literal text, `*` replaced by the current path field, `^` jump-forward or `^^` jump-backward segments, `?` conditionals, `|` alternatives) and tracks the positions it matched.
- **`xPath.vb`**: tokenizes a path string into segments. Each segment yields a prefix, a field name, an optional filter `[...]`, a sub-path `{{...}}`, or a multi-path list `{:a,b:}`. `Fetch()` steps through the segments.
- **`Parser.vb`**: the engine.
  - `xPath(path)` resolves a `{{sub-path}}` first, using a temporary `Parser` on the same buffer. It then walks the segments. Segment 0 searches the whole buffer, and each later segment searches inside the previous segment's results (a `Hashtable` of strings) through `FindSubText` / `findMultiSubText` → `ProcessToken`.
  - `filterOutput` applies the `=`, `!`, `~` filters.
  - `generateOutput` formats the results and applies the delete, keep and replace modes.
  - `addResultPath` and `addCleaningPath` support several `-x` and `-clean` paths. Cleaning paths are removed from the buffer before any query runs.
