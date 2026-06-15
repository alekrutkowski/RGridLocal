# RGridLocal

RGridLocal is an R package version of [RGrid](https://github.com/alekrutkowski/RGrid). It keeps the same browser spreadsheet interface and workbook/R-script format, but formulas are evaluated by the [standard R interpreter](https://www.r-project.org/) that starts the [localhost](https://simple.wikipedia.org/wiki/Localhost) server instead of by [webR](https://docs.r-wasm.org/).

## Install and open

```r
remotes::install_github("alekrutkowski/RGridLocal")
RGridLocal::openRGrid()
```

## Public API

The package exports one function:

```r
openRGrid(host = "127.0.0.1", port = 0L,
          launch.browser = interactive(),
          block = interactive(), quiet = FALSE)
```

`port = 0L` chooses an available random port. If you provide a positive port and it is unavailable, the next 50 ports are tried in order. Set `block = FALSE` to return a server handle immediately:

```r
srv <- RGridLocal::openRGrid(block = FALSE)
srv$url
srv$stop()
```

## Compatibility notes

The workbook format and exported RGrid R scripts remain compatible with [RGrid](https://github.com/alekrutkowski/RGrid). The local backend preserves the existing spreadsheet dependency graph, spill behavior, `ref()` semantics, named ranges and expressions, object summaries, plot pane, imports, and exports.

Missing packages detected in formulas are installed with `install.packages()` when they are not already available locally, matching the original app’s automatic package behavior but using the standard local R package library.

Plots are captured from a local PNG graphics device and returned to the browser as data URLs. Excel import/export still uses the existing SheetJS browser module, as in the original app.

## Security

Workbook formulas are executable R code in your local R session. Open workbooks and imported RGrid scripts only from sources you trust.
