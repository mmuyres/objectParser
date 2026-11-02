# Token maps

A token map tells the parser what the parts of a format look like. Each **prefix** you can use in
a path has one entry of three lines:

```text
prefix=<open token>
prefix=<close token>
prefix=<scope>
```

- The **open token** marks where a match starts.
- The **close token** marks where it ends.
- The **scope** says what to return:

| Scope | Returns | Example (XML map) |
|---|---|---|
| `I` | the text **inside**, between the tokens | `//-title` gives `Release Process` |
| `O` | the text **outside**, tokens included | `//title` gives `<title>Release Process</title>` |
| `F` | the **field**: the name `*` matched | `//variable/#` gives `flags`, `node`, … |
| `X\|Y` | per branch of a conditional close token (see below) | `O\|O` |

Other rules:

- The prefix is everything before the first `=`. A single space is the **default prefix**, used by
  path segments that have no prefix.
- Lines starting with `##` are comments. Blank lines are ignored.
- A later entry for the same prefix replaces an earlier one.
- `.` is reserved: it is the current-result step in paths.
- Load a map file with `-m file.map` (CLI) or `TokenMap.FromFile` (library). Problems are reported
  as warnings, never thrown.

## Token grammar

The parser classifies each token in this order:

| Form | Kind | Meaning |
|---|---|---|
| `%{}` `%{};` | balanced | from the current position, find the next `{` and go to the `}` that balances it. With a third character (`;`), give up if that character comes first. Any pair works: `%()`, `%[]`, `%''`. |
| `a?b\|c` | compare-conditional (close tokens) | if the open match was followed by `a` (or its jump target ended in `a`), the close token is `b`, otherwise `c` |
| `a\|b\|c` | search-conditional | find `a`, then the nearest of `b`, `c`, … |
| `a^^b` / `a^b` | jump | find `a`, then go forward to `b` |
| anything else | literal | find exactly this text |

Inside any token:

- `*` is replaced by the segment's **field** (its name). If the segment has no name, the field is
  read from the text (see [wildcards](path-syntax.md#empty-names-are-wildcards)).
- When `*` is at the start or end of the first part of a token, the match has to stop at a word
  boundary, so `<*` finds `<choice` but not `<choices`, and `*="` finds `name="` but not `rename="`.

### Escapes

Spaces at the end of a line disappear in many editors, and some characters are special. Write them
with a backslash:

| Escape | Means |
|---|---|
| `\n` `\r` `\t` | newline, carriage return, tab |
| `\s` | a space (use it for leading or trailing spaces) |
| `\\` | a backslash |
| `\*` `\^` `\|` `\?` `\%` | the character itself, not its special meaning |

`objParser -showmap` prints maps in this escaped form.

### Lookahead close tokens

A close token with nothing before its first `|` or `^` (for example `|\r|\n`) is a **lookahead**.
The match ends where the target was found, but the target isn't consumed, so the next match can
start with it. If none of the targets is found, the match runs to the end of the text.

That's what makes line-based formats work. Each bullet below ends at the `\n` that begins the next
one:

```text
li=\n\*\s
li=|\n|\r
li=I
```

### How XML elements are described

The default XML entry shows most of the grammar at once:

```text
 =<*^>
 =/>?/>|</*^>
 =O|O
```

1. Open `<*^>`: find `<` plus the name, then jump to the next `>`.
2. Close `/>?/>|</*^>`: if the open tag ended with `/>` it was self-closing, so the element is
   already complete. Otherwise find `</` plus the name, then jump to `>`.
3. Scope `O|O`: return the whole element in both cases.

## Worked example: MediaWiki markup

[`examples/mediawiki.map`](../examples/mediawiki.map) teaches the parser wiki markup in about 40
lines. Here is how each entry was arrived at.

**Headings.** `== Steps ==` sits on its own line. The open token is a newline, two equals signs
and a space. The close token is a space and two equals signs. The `\s` escapes keep the spaces
visible:

```text
h2=\n==\s
h2=\s==
h2=I
```

`//h2` gives `Steps`, `Owners`. Level-3 headings (`=== x ===`) don't match, because `\n==\s` needs
a space right after the second `=`.

**Link targets.** `[[Hotfix|hot fixes]]` should give `Hotfix`. Start at `[[` and stop at the
nearest `|` or `]]`. The `|` has to be escaped because it also separates the alternatives:

```text
target=[[
target=|\||]]
target=I
```

**Sections by name.** Make the section the default prefix, so `//Owners` works, and put `*` where
the heading text goes. The section ends at the next level-2 heading (lookahead) or at the end of
the page:

```text
 =\n==\s*\s==
 =|\n==\s|\n==\s
 =I
```

Now prefixes combine: `//Owners/li` gives the bullets of the Owners section only, and
`//li[.~'qa']` gives bullets that mention QA.

## More small maps

These are all covered by tests in `CustomMapTests.vb`.

Markdown links:

```text
text=[
text=]
text=I
url=](
url=)
url=I
```

INI-style sections and keys (`//db/@host`):

```text
 =\n[*]
 =|\n[|\n[
 =I
@=\n*=
@=|\r|\n
@=I
```

Log lines by level (`//ERROR`, `//{:WARN,ERROR:}`):

```text
 =[*]\s
 =|\r|\n
 =I
```

JSON-like objects and string values (`//user/@name`):

```text
@="*":\s"
@="
@=I
 ="*":\s
 =%{}
 =O
```

## Tips

- Start from a built-in map: `objParser -showmap -xml > my.map`.
- Run with `-debug` to see each token search and where it matched (written to stderr).
- Test a single entry quickly with ad-hoc delimiters: `objParser -f file -d "<open>" -D "<close>"`
  (`~` stands for a double quote there).
