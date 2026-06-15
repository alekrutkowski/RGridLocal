# Contributing

RGrid-local is served by the R package entry point `openRGrid()`. The browser assets live under `inst/app`; the local R server and API live in `R/openRGrid.R`.

Useful checks from the package root:

```sh
node --check inst/app/app.js
python -m py_compile inst/app/serve.py
```

A live functional check requires an installed R interpreter with `httpuv` and `jsonlite`.
