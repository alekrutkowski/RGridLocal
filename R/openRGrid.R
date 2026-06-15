`%||%` <- function(x, y) if (is.null(x) || !length(x)) y else x

#' Open RGrid-local
#'
#' Start a localhost RGrid session whose spreadsheet formulas are evaluated by
#' the standard R interpreter that is running this package.
#'
#' @param host Interface to bind. The default only accepts connections from the
#'   local machine.
#' @param port TCP port to try first. Use `0` or `NULL` to choose an
#'   available random port. If a positive port is unavailable, the next 50
#'   ports are tried in order.
#' @param launch.browser Logical or function. When `TRUE`, open the browser
#'   after the server starts. When a function, it is called with the URL.
#' @param block Keep servicing requests until the user interrupts R. Set to
#'   `FALSE` to return a server handle immediately.
#' @param quiet Suppress startup messages.
#'
#' @return Invisibly returns a server handle when `block = FALSE`; otherwise the
#'   server runs until interrupted.
#' @export
#'
#' @examples
#' \dontrun{
#' openRGrid()
#' }
openRGrid <- function(host = "127.0.0.1", port = 0L,
                      launch.browser = interactive(), block = interactive(), quiet = FALSE) {
  if (!requireNamespace("httpuv", quietly = TRUE)) {
    stop("RGrid-local requires the httpuv package. Install it with install.packages('httpuv').", call. = FALSE)
  }
  if (!requireNamespace("jsonlite", quietly = TRUE)) {
    stop("RGrid-local requires the jsonlite package. Install it with install.packages('jsonlite').", call. = FALSE)
  }
  app_dir <- rgrid_app_dir()
  app <- rgrid_http_app(app_dir)

  if (is.null(port) || identical(suppressWarnings(as.integer(port)[1L]), 0L)) {
    requested_port <- httpuv::randomPort()
  } else {
    requested_port <- suppressWarnings(as.integer(port)[1L])
  }
  if (is.na(requested_port) || requested_port < 1L || requested_port > 65535L) {
    stop("port must be an integer between 1 and 65535, 0, or NULL.", call. = FALSE)
  }

  server <- NULL
  actual_port <- NA_integer_
  last_error <- NULL
  for (candidate in seq.int(requested_port, min(65535L, requested_port + 50L))) {
    server <- tryCatch(httpuv::startServer(host, candidate, app), error = function(e) {
      last_error <<- e
      NULL
    })
    if (!is.null(server)) {
      actual_port <- candidate
      break
    }
  }
  if (is.null(server)) {
    stop("Could not start RGrid-local near port ", requested_port, ": ",
         conditionMessage(last_error), call. = FALSE)
  }

  on.exit(try(httpuv::stopServer(server), silent = TRUE), add = TRUE)
  display_host <- if (identical(host, "0.0.0.0") || identical(host, "::")) "127.0.0.1" else host
  url <- sprintf("http://%s:%d/", display_host, actual_port)

  if (!quiet) {
    message("RGrid-local: ", url)
    if (isTRUE(block)) message("Close this session with Esc or Ctrl+C in the R console.")
  }
  if (is.function(launch.browser)) {
    launch.browser(url)
  } else if (isTRUE(launch.browser)) {
    utils::browseURL(url)
  }

  handle <- structure(
    list(
      url = url,
      host = host,
      port = actual_port,
      server = server,
      stop = function() httpuv::stopServer(server)
    ),
    class = "RGridLocalServer"
  )

  if (!isTRUE(block)) {
    on.exit(NULL, add = FALSE)
    return(invisible(handle))
  }

  repeat {
    httpuv::service(100)
  }

  invisible(handle)
}

rgrid_app_dir <- function() {
  installed <- system.file("app", package = "RGridLocal", mustWork = FALSE)
  if (nzchar(installed) && dir.exists(installed)) return(installed)
  candidates <- c(
    file.path(getwd(), "inst", "app"),
    file.path(dirname(getwd()), "RGrid-local", "inst", "app")
  )
  for (candidate in candidates) {
    if (dir.exists(candidate)) return(normalizePath(candidate, winslash = "/", mustWork = TRUE))
  }
  stop("RGrid-local app files were not found. Install the package or run from the package source root.", call. = FALSE)
}

rgrid_http_app <- function(app_dir) {
  force(app_dir)
  list(
    call = function(req) {
      path <- req$PATH_INFO %||% req$REQUEST_URI %||% "/"
      path <- sub("\\?.*$", "", path)
      if (startsWith(path, "/__rgrid_api")) {
        return(rgrid_handle_api(req, path))
      }
      rgrid_static_response(app_dir, path)
    }
  )
}

rgrid_handle_api <- function(req, path) {
  endpoint <- sub("^/__rgrid_api/?", "", path)
  tryCatch({
    if (endpoint %in% c("", "status")) {
      return(rgrid_json_response(list(
        ok = TRUE,
        r_version = as.character(getRversion()),
        platform = R.version$platform
      )))
    }

    payload <- rgrid_read_json(req)
    code <- as.character(payload$code %||% "")

    switch(endpoint,
      "eval-void" = {
        eval(parse(text = code), envir = .GlobalEnv)
        rgrid_json_response(list(ok = TRUE, value = NULL))
      },
      "eval-boolean" = {
        value <- eval(parse(text = code), envir = .GlobalEnv)
        rgrid_json_response(list(ok = TRUE, value = isTRUE(value)))
      },
      "eval-string" = {
        value <- eval(parse(text = code), envir = .GlobalEnv)
        text <- if (is.null(value) || !length(value)) "" else paste(as.character(value), collapse = "\n")
        rgrid_json_response(list(ok = TRUE, value = text))
      },
      "install-packages" = {
        packages <- unique(as.character(unlist(payload$packages %||% character(), use.names = FALSE)))
        rgrid_install_packages(packages)
        rgrid_json_response(list(ok = TRUE, packages = as.list(packages)))
      },
      "capture" = {
        options <- payload$options %||% list()
        if (!is.null(payload$plot)) options$captureGraphics <- payload$plot
        result <- rgrid_capture_eval(code, options)
        rgrid_json_response(c(list(ok = TRUE), result))
      },
      rgrid_json_response(list(ok = FALSE, error = paste("Unknown RGrid-local API endpoint:", endpoint)), status = 404L)
    )
  }, error = function(e) {
    rgrid_json_response(list(ok = FALSE, error = conditionMessage(e)), status = 500L)
  })
}

rgrid_install_packages <- function(packages) {
  packages <- packages[nzchar(packages)]
  if (!length(packages)) return(invisible(character()))
  missing <- packages[!vapply(packages, requireNamespace, logical(1L), quietly = TRUE)]
  if (!length(missing)) return(invisible(character()))
  repositories <- getOption("repos")
  if (!length(repositories) || identical(unname(repositories["CRAN"]), "@CRAN@")) {
    repositories <- c(CRAN = "https://cloud.r-project.org")
  }
  install.packages(missing, repos = repositories)
  still_missing <- missing[!vapply(missing, requireNamespace, logical(1L), quietly = TRUE)]
  if (length(still_missing)) {
    stop("Packages remain unavailable: ", paste(still_missing, collapse = ", "), call. = FALSE)
  }
  invisible(missing)
}

rgrid_capture_eval <- function(code, options) {
  graphics_options <- options$captureGraphics %||% FALSE
  if (!is.list(graphics_options)) graphics_options <- list(capture = FALSE)
  capture_graphics <- isTRUE(graphics_options$capture)
  width <- as.integer(graphics_options$width %||% 1008L)
  height <- as.integer(graphics_options$height %||% 672L)
  bg <- as.character(graphics_options$bg %||% "white")
  width <- max(1L, min(4096L, width))
  height <- max(1L, min(4096L, height))

  temp_dir <- tempfile("rgrid-plots-")
  dir.create(temp_dir, recursive = TRUE, showWarnings = FALSE)
  on.exit(unlink(temp_dir, recursive = TRUE), add = TRUE)

  opened_device <- NULL
  if (capture_graphics && isTRUE(capabilities("png"))) {
    grDevices::png(filename = file.path(temp_dir, "plot-%03d.png"), width = width, height = height, bg = bg)
    opened_device <- grDevices::dev.cur()
  }
  on.exit(rgrid_close_device(opened_device), add = TRUE)

  captured_conditions <- character()
  captured_output <- withCallingHandlers(
    utils::capture.output(
      result <- eval(parse(text = code), envir = .GlobalEnv),
      type = "output"
    ),
    message = function(m) {
      captured_conditions <<- c(captured_conditions, conditionMessage(m))
      invokeRestart("muffleMessage")
    },
    warning = function(w) {
      captured_conditions <<- c(captured_conditions, conditionMessage(w))
      invokeRestart("muffleWarning")
    }
  )
  value <- if (exists("result", inherits = FALSE)) result else NULL

  rgrid_close_device(opened_device)
  opened_device <- NULL

  list(
    result = rgrid_encode_packed_result(value),
    images = rgrid_collect_plot_images(temp_dir, width, height, bg),
    output = as.list(c(captured_output, captured_conditions))
  )
}

rgrid_close_device <- function(device) {
  if (is.null(device) || is.na(device)) return(invisible(FALSE))
  devices <- grDevices::dev.list()
  if (is.null(devices) || !(device %in% devices)) return(invisible(FALSE))
  try(grDevices::dev.set(device), silent = TRUE)
  try(grDevices::dev.off(), silent = TRUE)
  invisible(TRUE)
}

rgrid_collect_plot_images <- function(temp_dir, width, height, bg) {
  files <- sort(list.files(temp_dir, pattern = "^plot-[0-9]+[.]png$", full.names = TRUE))
  if (!length(files)) return(list())
  blank <- rgrid_blank_png(width, height, bg)
  images <- lapply(files, function(file) {
    info <- file.info(file)
    if (is.na(info$size) || info$size <= 0) return(NULL)
    bytes <- readBin(file, what = "raw", n = info$size)
    if (identical(bytes, blank)) return(NULL)
    if (length(bytes) < 5000L && length(blank) < 5000L && abs(length(bytes) - length(blank)) < 64L) return(NULL)
    list(width = width, height = height, dataUrl = paste0("data:image/png;base64,", jsonlite::base64_enc(bytes)))
  })
  Filter(Negate(is.null), images)
}

rgrid_blank_png <- function(width, height, bg) {
  file <- tempfile(fileext = ".png")
  on.exit(unlink(file), add = TRUE)
  grDevices::png(filename = file, width = width, height = height, bg = bg)
  grDevices::dev.off()
  info <- file.info(file)
  if (is.na(info$size) || info$size <= 0) return(raw())
  readBin(file, what = "raw", n = info$size)
}

rgrid_encode_packed_result <- function(value) {
  if (!is.list(value) || is.null(value$ok)) {
    return(list(ok = FALSE, error = "RGrid-local expected a packed RGrid result but received a different object."))
  }
  if (!isTRUE(value$ok)) {
    return(list(ok = FALSE, error = as.character(value$error %||% "R evaluation failed")))
  }
  labels <- as.character(value$tree_label %||% character())
  types <- as.character(value$tree_type %||% character())
  summaries <- as.character(value$tree_summary %||% character())
  parents <- as.integer(value$tree_parent %||% integer())
  n_tree <- max(length(labels), length(types), length(summaries), length(parents), 0L)
  tree <- if (n_tree) lapply(seq_len(n_tree), function(i) {
    list(
      index = i,
      label = rgrid_chr_at(labels, i),
      type = rgrid_chr_at(types, i),
      summary = rgrid_chr_at(summaries, i),
      parent = rgrid_int_at(parents, i)
    )
  }) else list()

  list(
    ok = TRUE,
    nrow = as.integer(value$nrow %||% 1L),
    ncol = as.integer(value$ncol %||% 1L),
    kind = as.character(value$kind %||% ""),
    object_kind = as.character(value$object_kind %||% "other"),
    plot_kind = as.character(value$plot_kind %||% ""),
    object_class = as.character(value$object_class %||% ""),
    preserve_object = isTRUE(value$preserve_object),
    values = lapply(as.vector(value$values), rgrid_json_value),
    tree = tree
  )
}

rgrid_chr_at <- function(x, i) {
  if (length(x) < i || is.na(x[i])) "" else enc2utf8(as.character(x[i]))
}

rgrid_int_at <- function(x, i) {
  if (length(x) < i || is.na(x[i])) 0L else as.integer(x[i])
}

rgrid_json_value <- function(x) {
  if (is.null(x) || !length(x)) return(NULL)
  x <- x[[1L]]
  if (is.factor(x)) x <- as.character(x)
  if (inherits(x, c("Date", "POSIXct", "POSIXlt", "POSIXt"))) x <- as.character(x)
  if (is.raw(x)) x <- as.integer(x)
  if (is.complex(x)) x <- as.character(x)
  if (is.logical(x)) {
    if (is.na(x)) return(NULL)
    return(isTRUE(x))
  }
  if (is.numeric(x)) {
    if (is.nan(x)) return("NaN")
    if (is.na(x)) return(NULL)
    if (!is.finite(x)) return(as.character(x))
    return(unname(as.numeric(x)))
  }
  if (is.character(x)) {
    if (is.na(x)) return(NULL)
    return(enc2utf8(x))
  }
  as.character(x)
}

rgrid_read_json <- function(req) {
  body <- rgrid_request_body(req)
  if (!nzchar(body)) return(list())
  jsonlite::fromJSON(body, simplifyVector = FALSE)
}

rgrid_request_body <- function(req) {
  input <- req[["rook.input"]]
  if (is.null(input) || is.null(input$read)) return("")
  body <- input$read()
  if (is.raw(body)) rawToChar(body) else paste(body, collapse = "")
}

rgrid_json_response <- function(value, status = 200L) {
  rgrid_response(
    jsonlite::toJSON(value, auto_unbox = TRUE, null = "null", na = "null", digits = NA),
    status = status,
    type = "application/json; charset=UTF-8"
  )
}

rgrid_static_response <- function(app_dir, path) {
  path <- utils::URLdecode(path)
  if (identical(path, "/") || identical(path, "")) path <- "/index.html"
  path <- sub("^/+", "", path)
  if (!nzchar(path) || grepl("(^|/)[.][.](/|$)|[\\\\]", path)) {
    return(rgrid_response("Not found", status = 404L, type = "text/plain; charset=UTF-8"))
  }
  root <- normalizePath(app_dir, winslash = "/", mustWork = TRUE)
  file <- normalizePath(file.path(app_dir, path), winslash = "/", mustWork = FALSE)
  prefix <- paste0(root, "/")
  if (!startsWith(tolower(file), tolower(prefix)) && !identical(tolower(file), tolower(root))) {
    return(rgrid_response("Not found", status = 404L, type = "text/plain; charset=UTF-8"))
  }
  if (!file.exists(file) || dir.exists(file)) {
    return(rgrid_response("Not found", status = 404L, type = "text/plain; charset=UTF-8"))
  }
  info <- file.info(file)
  bytes <- readBin(file, what = "raw", n = info$size)
  rgrid_response(bytes, status = 200L, type = rgrid_mime_type(file))
}

rgrid_response <- function(body, status = 200L, type = "text/plain; charset=UTF-8") {
  list(
    status = as.integer(status),
    headers = list(
      "Content-Type" = type,
      "Cache-Control" = "no-cache",
      "X-Content-Type-Options" = "nosniff"
    ),
    body = body
  )
}

rgrid_mime_type <- function(file) {
  ext <- tolower(tools::file_ext(file))
  switch(ext,
    html = "text/html; charset=UTF-8",
    htm = "text/html; charset=UTF-8",
    js = "text/javascript; charset=UTF-8",
    mjs = "text/javascript; charset=UTF-8",
    css = "text/css; charset=UTF-8",
    svg = "image/svg+xml",
    png = "image/png",
    json = "application/json; charset=UTF-8",
    webmanifest = "application/manifest+json; charset=UTF-8",
    txt = "text/plain; charset=UTF-8",
    "application/octet-stream"
  )
}
