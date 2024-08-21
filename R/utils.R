# replacement for stopifnot() with nicer user feedback
stop_if_not_class <- \(x, class) {
  check <- inherits(x, what = class)
  if (!check) {
    cli::cli_abort("Expected class {.val {class}}. Got {.val {class(x)}}")
  }
}


stop_if_not_rpptx <- \(x) {
  stop_if_not_class(x, "rpptx")
}


stop_if_not_robopptx <- \(x) {
  stop_if_not_class(x, "robopptx")
}


#' Open local file in default application
#' @param path Path to file.
#' @export
#' @examples \dontrun{
#' file <- system.file("img", "Rlogo.png", package = "png")
#' file_open(file)
#' }
#'
file_open <- function(path) {
  path <- normalizePath(path)
  path_quoted <- shQuote(path)
  os_type <- .Platform$OS.type
  if (os_type == "windows") {
    shell.exec(path)
  } else {
    if (Sys.info()["sysname"] == "Darwin") {
      system(paste("open", path_quoted))
    } else {
      system(paste("xdg-open", path_quoted))
    }
  }
}
