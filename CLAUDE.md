# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

objParser queries structured text (XML, HTML, wiki markup, VB.NET, C#, anything delimiter-based)
with a pseudo-XPath, **without building a tree**. A token map assigns each path prefix an
open/close token pair. Each path segment string-searches for its pair inside the results of the
previous segment. User docs are in `README.md` and `docs/`; read `docs/architecture.md` before
changing the engine.

## Commands

```sh
dotnet build ObjectParser.sln
dotnet test ObjectParser.sln                                  # ~517 tests, a few seconds
dotnet test ObjectParser.sln --filter "FullyQualifiedName~FilterTests"            # one class
dotnet test ObjectParser.sln --filter "FullyQualifiedName~FilterTests.Equals_literal"  # one test
dotnet run --project src/objParser -- -f examples/model.xml -x "//variable/@name"
dotnet publish src/objParser -p:PublishProfile=win-x64        # free-standing exe → dist/win-x64/objParser.exe (checked in: commit it with source changes)
```

The publish profile (`src/objParser/Properties/PublishProfiles/win-x64.pubxml`) produces one
self-contained, trimmed, single-file exe (~11 MB, no .NET install needed). Trimming is warning-free
because everything is `Option Strict On` with no late binding. Keep it that way: late binding or
reflection in the library can break the trimmed exe while every test still passes. After changing
the CLI, publish and run the exe once.

.NET 9, VB.NET, `Option Strict On`, `TreatWarningsAsErrors` (in `Directory.Build.props`). In Git
Bash, run the exe with `MSYS_NO_PATHCONV=1`, or arguments like `</node>` get rewritten.

## Layout

- `src/ObjectParser/`: the library.
  - `Parser.vb`: evaluation loop (`Evaluate` → `FindInText` → `ProcessToken`), filters, wildcard field derivation, balanced matching, `Delete`.
  - `ParserToken.vb`: token grammar and match positions.
  - `TokenMap.vb`: map-file parsing, escapes, `CreatePair`.
  - `PathExpression.vb`: path → segments.
  - `BuiltInMaps.vb`: the xml / xml_ext / vb / cs maps, written as map text.
- `src/objParser/`: the CLI. All logic is in `Cli.Run(args, in, out, err)` so tests call it in-process.
- `tests/ObjectParser.Tests/`: xUnit, in VB. `CrossCheckTests` compares 300 generated XML documents with what the generator wrote and with System.Xml.Linq; `RobustnessTests` feeds random input with a timeout.
- `examples/`: `model.xml`, `mediawiki.map`, `release.wiki`. The README quotes them, and the test project links them in as TestData, so keep the README output in sync with the tests.

## Engine rules that are easy to break

- `TokenMap.CreatePair` returns a **fresh** `TokenPair` every time. Never cache or share pairs: filters and sub-paths run nested `Parser`s on the same map.
- Results are `Match(start, value)` with absolute offsets. `Delete` splices by position, and tests assert `Text.Substring(m.Start, m.Length) = m.Value` everywhere.
- A token's flags are re-parsed by `Reset()`/`Init()`, but `Role` (Open/Close) and `LookaheadSize` survive. The close token is reset on every loop iteration in `FindInText`.
- A close token with an empty first tag (`|a|b`, `^a`) is a lookahead (`ParserToken.IsLookahead`): it doesn't consume its target, and for a search-conditional, end of text counts as a match.
- Scope `F` returns the field. Wildcard derivation (empty or `*` field) reads name characters after the token's lead literal, or before its trail literal. A map defining `*` as a prefix disables wildcards.
- After the first segment, the search starts past the parent's opening tag when the parent pair's IO is `O`. So `//li/-li` is empty by design, and `.` re-enters a block from its start.
- Map text escapes (`\*`, `\|` …) become private-use placeholder characters (`TokenMap.LiteralBase`). `TokenMap.Literal()` must be applied wherever token text is compared with input text.

## Gotchas when editing

- VB is case-insensitive: a local `q` or `m` collides with the test helpers `Q`/`M`.
- Significant spaces at the end of a map line get stripped by editors. Write them as `\s`.
- The original 2010 sources are under the git tag `v4.1-legacy`. `docs/changes.md` lists every intentional behaviour change from v4.1, so update it when you change behaviour.
