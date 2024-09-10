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

stop_if_not_correct_shapeid <- \(shapeid, robopptx, robo_content = NA_character_) {
  stop_if_not_robopptx(robopptx)
  if (is.na(robo_content)) {
    possible_ids <- robopptx$robocop$get_layout_overview()[, sort(unique(id))]
  } else {
    filter_content <- robo_content
    possible_ids <- robopptx$robocop$get_layout_overview()[robo_content == filter_content, sort(unique(id))]
  }
  check <- all(is.na(shapeid)) || all(as.character(shapeid) %chin% possible_ids)

  if (!check) {
    cli::cli_abort(
      "Expected shapeids {.val {possible_ids}}. Got {.val {sort(setdiff(shapeid, possible_ids))}}"
    )
  }
}

stop_if_not_correct_layoutid <- \(layoutid, robopptx, robo_content = NA_character_) {
  stop_if_not_robopptx(robopptx)
  if (is.na(robo_content)) {
    possible_layoutids <- robopptx$robocop$get_layout_overview()[, sort(unique(layout_id))]
  } else {
    filter_content <- robo_content
    possible_layoutids <- robopptx$robocop$get_layout_overview()[robo_content == filter_content, sort(unique(layout_id))]
  }
  check <- all(is.na(layoutid)) || all(layoutid %in% possible_layoutids)

  if (!check) {
    cli::cli_abort(
      "Expected shapeids {.val {possible_layoutids}}. Got {.val {sort(setdiff(layoutid, possible_layoutids))}}"
    )
  }
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
