# Path syntax

A path is a list of **segments** separated by `/`. Each segment is evaluated inside every result of
the segment before it.

```text
//projatt/att[option]/.[@type='pulldown']/@name
  └──┬──┘ └─┬─┘└──┬─┘ │└───────┬────────┘ └─┬─┘
  segment  segment filter│    filter        segment (prefix @, name "name")
                current result
```

A leading `//` or `/` is optional and makes no difference: every path starts by searching the
whole text.

## Segments: prefix + name

A segment is an optional **prefix** followed by a **name**. The prefix chooses a token pair from
the map, and the name is substituted for `*` in that pair's tokens.

| Segment | Prefix | Name | With the XML map, finds |
|---|---|---|---|
| `variable` | (default, a space) | `variable` | `<variable …>…</variable>` |
| `-node` | `-` | `node` | the text inside `<node>…</node>` |
| `@name` | `@` | `name` | the value of `name="…"` |
| `@` | `@` | (none) | every attribute value |
| `$` | `$` | (none) | every attribute name |

The longest matching prefix wins, so in the `xml_ext` map `nodename:` beats `node:`. A segment
whose prefix isn't in the map is skipped.

### Empty names are wildcards

If a segment has no name (or the name is `*`) and the token contains `*`, the name is **read from
the text**: the run of name characters right after the token's leading text, or right before its
trailing text when there's nothing in front.

```text
//projatt/          direct children of each projatt element
//*                 the top-level element
//Invoice/!         (vb) every function in class Invoice
//variable/@        every attribute value inside each variable
```

Name characters are letters, digits, `_`, `-`, `.` and `:`. That's also why `//choice` doesn't
match `<choices>`: a name has to stop at a non-name character.

### Searches go through the whole block

A segment searches *everything* inside the previous result, not only direct children.
`//model/@name` returns the `name` of `model` and of every element inside it.

After the first segment, the search starts just past the parent's opening tag, so the parent isn't
found again. A consequence: `//li/-li` finds nothing. Use `//-li` for the content of every `li`,
or `//li/!` for the first run of text in each one (handy after a filter: `//li[@class~'done']/!`).

Elements nested inside an element *of the same name* (`<a><a/></a>`) are not supported: the
close token stops at the first `</a>`.

## Filters `[...]`

A filter after a segment keeps only the results that pass. The **left** side is a path evaluated
inside each result. The **right** side is one of:

- a literal in single quotes: `'number'`
- a path starting with `//`, evaluated over the whole text
- any other path, evaluated inside the *parent* of the current results

| Filter | Keeps a result when |
|---|---|
| `[x]` | `x` finds something |
| `[x!]` or `[!x]` | `x` finds nothing |
| `[x='v']` | any value of `x` equals `v` |
| `[x!'v']` | no value of `x` equals `v` |
| `[x~'v']` | any value of `x` contains `v` |
| `[x=//y]` | any value of `x` equals any value of `//y` |
| `[x!//y]` | no value of `x` equals any value of `//y` |
| `[x~//y]` | any value of `x` contains any value of `//y` |

Comparisons ignore case and line breaks. Quotes protect operator characters, so
`[@k='a=b']` works.

Examples against [`examples/model.xml`](../examples/model.xml):

```text
//variable[@type='number']/@name                     alpha gamma
//variable[!flags]/@name                             gamma
//variable[@name~'ph']/@name                         alpha
//choices/choice[@type=//controls/type/@value]       choices whose type is a known control
//g/item[@name=ref/@to]                              items named by a ref in the same g
//table/tr[-td='QA']                                 (html) rows with a cell whose text is QA
```

`[+x]` is accepted for compatibility with v4.1 and means `[x]`. A leading `-` is an ordinary
prefix (in the XML maps, "inner text").

### Several filters: the current result `.`

`.` passes each result through unchanged, so you can add another filter:

```text
//projatt/att/option[@id='7']/.[@name='CNS-989']
//projatt/att[option!]/.[@snapshot_output]/@name
```

Change the token with `-results_token <t>` (CLI) or `ParserSyntax.CurrentResultToken`.

## Sub-paths `{{...}}`

One sub-path may appear in a path. It is evaluated first, over the whole text, and its results
replace it:

```text
//variable[@name='{{//index[flags/flag/@attr='portfolio_index']/@name}}']
```

If the sub-path returns several values inside a quoted literal, the literal acts as a list (any of
them can match).

## Multi-paths `{:...:}`

A comma-separated list of names in one segment. Results come name by name:

```text
//{:variable,decision,objective:}/@name      alpha beta gamma d1 o1
```

`{:*:}` repeats the most recent list. Here, with the `xml_ext` map's `node:` prefix, it returns
the opening tag of each element that has flags. The `.` step is what lets the search see the
element's own opening tag again:

```text
//{:variable,decision,objective:}[flags]/./node:{:*:}
    <variable name="alpha" type="number">
    <variable name="beta" type="text">
    <decision name="d1">
    <objective name="o1">
```

## Results

Results are returned in document order, each with its position in the text. The CLI prints one
per line, `-s` sorts them and removes duplicates, and `-count` reports how many there were.
