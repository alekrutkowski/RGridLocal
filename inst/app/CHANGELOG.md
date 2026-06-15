# Changelog

## RGrid-local 0.1.0

- Converted the browser app into an R package with one exported function, `openRGrid()`.
- Replaced webR with a localhost API served by the R session that called `openRGrid()`.
- Preserved the RGrid workbook and exported R-script formats.
- Preserved the active row/column header highlight and formula-editor double-click height toggle from the supplied codebase.
