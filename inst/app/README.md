# RGrid-local

RGrid-local is the R package version of RGrid. It keeps the same browser spreadsheet interface and workbook/R-script format, but formulas are evaluated by the standard R interpreter that starts the localhost server instead of by webR.

## Install and open

From the package source directory or from the zip after extracting it:

```r
install.packages(c("httpuv", "jsonlite"))
install.packages("RGrid-local", repos = NULL, type = "source")
RGridLocal::openRGrid()
```

For development without installing the package, source the function and run it from the package root:

```r
source("R/openRGrid.R")
openRGrid()
```

`openRGrid()` serves the app on `127.0.0.1` and opens it in your browser. The R session remains the calculation backend while the server is running. Interrupt R to stop a blocking server.

Use a non-blocking server handle like this:

```r
srv <- RGridLocal::openRGrid(block = FALSE)
srv$url
srv$stop()
```

## Public API

The package exports one function:

```r
openRGrid(host = "127.0.0.1", port = 0L,
          launch.browser = interactive(), block = interactive(), quiet = FALSE)
```

## Compatibility notes

The workbook format and exported RGrid R scripts remain compatible with RGrid. The local backend preserves the existing spreadsheet dependency graph, spill behavior, `ref()` semantics, named ranges and expressions, object summaries, plot pane, imports, and exports.

Plots are captured from a local PNG graphics device and returned to the browser as data URLs. Excel import/export still uses the existing SheetJS browser module, as in the original app.

Missing R packages referenced by formulas are installed into the local R library using `install.packages()`, matching the original app's automatic package-resolution behavior but in the standard local R interpreter.

## Security

Workbook formulas are executable R code in your local R session. Open workbooks and imported RGrid scripts only from sources you trust.
