# Run with: Rscript serve.R
# Requires install.packages("httpuv") once in the system R installation.

library(httpuv)

args <- commandArgs(trailingOnly = FALSE)
file_arg <- sub("^--file=", "", args[grep("^--file=", args)][1])
root <- dirname(normalizePath(file_arg, mustWork = TRUE))
setwd(root)
cat("RGrid: http://127.0.0.1:8080\n")

runServer(
  host = "127.0.0.1",
  port = 8080,
  app = list(
    staticPaths = list(
      "/" = staticPath(
        ".",
        headers = list(
          "Cross-Origin-Opener-Policy" = "same-origin",
          "Cross-Origin-Embedder-Policy" = "require-corp",
          "Cross-Origin-Resource-Policy" = "cross-origin",
          "Cache-Control" = "no-cache"
        )
      )
    )
  )
)
