# Security policy

RGrid-local executes workbook formulas as R code in the local R session that started `openRGrid()`. Workbooks and imported RGrid scripts should therefore be treated as executable code.

Only open files from sources you trust. Keep the server bound to `127.0.0.1` unless you have a specific reason to expose it on another network interface.
