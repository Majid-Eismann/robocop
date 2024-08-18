#' Add content to slide candidate
#'
#' @param robopptx robopptx Imported layout file from \link{load_layout}
#' @param content object Content to add to the current slide candidate
#'
#' @return
#' @export
#'
#' @examples
add <- function(robopptx, content = NULL, robo_class = NA_character_, r_class = NA_character_) {
  stopifnot("robopptx" %in% class(my_layout))

  add_position <- robopptx$robocop$slide_candidate[!is.na(add_order), .N] + 1

  robopptx$robocop$slide_candidate[
    add_position,
    ":="(
      add_order = add_position,
      robo_class = ifelse(
        is.na(..robo_class),
        NA_character_,
        ..robo_class
      ),
      class_r = ifelse(
        is.na(r_class),
        class(..content)[length(class(..content))],
        r_class
      ),
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
  add(robopptx = robopptx, content = title, robo_class = "title")
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
  add(robopptx = robopptx, content = subtitle, robo_class = "subtitle")
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
  add(robopptx = robopptx, content = text, robo_class = "text")
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
  add(robopptx = robopptx, content = footer, robo_class = "footer")
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

  add(robopptx = robopptx, content = table, robo_class = "table")
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
  if (class(graph)[1] == "character") {
    stopifnot(file.exists(graph))
    stopifnot(grepl("[.](img|png)$", graph))

    add(robopptx, content = graph, robo_class = "graph", r_class = "external_img")

  } else {
    add(robopptx, content = graph, robo_class = "graph")
  }

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

  match.arg(names(args), unique(robopptx$robocop$class_mapping$robo_class), several.ok = TRUE)

  robopptx$robocop$add_slide(args)

  invisible(robopptx)
}
