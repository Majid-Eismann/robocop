#' Add content to slide candidate
#'
#' @param robopptx robopptx Imported layout file from \link{load_layout}
#' @param content object Content to add to the current slide candidate
#'
#' @return
#' @export
#'
#' @examples
add <- function(robopptx, content = NULL, class_easy = NA_character_) {
  stopifnot("robopptx" %in% class(my_layout))

  add_position <- robopptx$robocop$slide_candidate[!is.na(add_order), .N] + 1

  robopptx$robocop$slide_candidate[
    add_position,
    ":="(
      add_order = add_position,
      class_easy = ..class_easy,
      class_r = group_r_class(..content),
      content = list(list(..content))
    )
  ]

  invisible(robopptx)
}

#' Add title to slide candidate
#'
#' @param robopptx robopptx Imported layout file from \link{load_layout}
#' @param title character Title to add to the current slide candidate
#'
#' @return
#' @export
#'
#' @examples
add_title <- function(robopptx, title = NULL) {
  add(robopptx = robopptx, content = title, class_easy = "title")
}

#' Add subtitle to slide candidate
#'
#' @param robopptx robopptx Imported layout file from \link{load_layout}
#' @param sub_title character Subtitle to add to the current slide candidate
#'
#' @return
#' @export
#'
#' @examples
add_subtitle <- function(robopptx, subtitle = NULL) {
  add(robopptx = robopptx, content = subtitle, class_easy = "subtitle")
}

#' Add Text to slide candidate
#'
#' @param robopptx robopptx Imported layout file from \link{load_layout}
#' @param text character Text to add to the current slide candidate
#'
#' @return
#' @export
#'
#' @examples
add_text <- function(robopptx, text = NULL) {
  add(robopptx = robopptx, content = text, class_easy = "text")
}

#' Add Text to slide candidate
#'
#' @param robopptx robopptx Imported layout file from \link{load_layout}
#' @param footer character Footer to add to the current slide candidate
#'
#' @return
#' @export
#'
#' @examples
add_footer <- function(robopptx, footer = NULL) {
  add(robopptx = robopptx, content = footer, class_easy = "footer")
}

#' Add Table to slide candidate
#'
#' @param robopptx robopptx Imported layout file from \link{load_layout}
#' @param table data.frame Table to add to the current slide candidate
#'
#' @return
#' @export
#'
#' @examples
add_table <- function(robopptx, table = NULL) {
  table <- as.data.frame(table)

  add(robopptx = robopptx, content = table, class_easy = "table")
}

#' Add Graph to slide candidate
#'
#' @param robopptx robopptx Imported layout file from \link{load_layout}
#' @param graph ggplot, mschart, charater Graph to add to the current slide candidate as ggplot, mschart or filepath to existing .img or .png file
#'
#' @return
#' @export
#'
#' @examples
add_graph <- function(robopptx, graph = NULL) {
  # if image is given as path - check file and filetype
  if (group_r_class(graph) == "character") {
    stopifnot(file.exists(graph))
    stopifnot(grepl("[.](img|png)$", graph))
  }

  add(robopptx, content = graph, class_easy = "graph")
}

#' Add content to slide candidate
#'
#' @param robopptx robopptx Imported layout file from \link{load_layout}
#' @param ... object Content to add to the current slide candidate
#'
#' @return
#' @export
#'
#' @examples
add_slide <- function(robopptx, ...) {
  stopifnot("robopptx" %in% class(robopptx))

  args <- list(...)

  if (is.null(names(args)) && !is.null(names(unlist(args, recursive = F)))) {
    args <- unlist(args, recursive = F)
  }

  match.arg(names(args), unique(robopptx$robocop$class_mapping$class_easy), several.ok = TRUE)

  robopptx$robocop$add_slide(args)

  invisible(robopptx)
}
