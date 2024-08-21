#' Add content to slide candidate
#'
#' @param robopptx robopptx Imported layout file from \link{load_layout}
#' @param content object Content to add to the current slide candidate
#'
#' @return A `robopptx` object.
#' @export
#'
#' @examples "todo"
add <- function(robopptx, content = NULL, slide = NA_integer_, shapeid = NA_integer_,
                hint = NULL, ..., robo_class = NA_character_, r_class = NA_character_) {
  # stopifnot("robopptx" %in% class(my_layout))
  stop_if_not_robopptx(my_layout)

  dots <- list(...)

  add_position <- robopptx$robocop$slide_candidate[!is.na(add_order), .N] + 1

  robopptx$robocop$slide_candidate[
    add_position,
    ":="(
      add_order = add_position,
      content = list(list(..content)),
      user_selection_slide = slide,
      user_selection_shapeid = shapeid,
      hint = list(list(hint)),
      dotdotdot = list(list(dots)),
      robo_class = ifelse(
        is.na(..robo_class),
        NA_character_,
        ..robo_class
      ),
      class_r = ifelse(
        is.na(r_class),
        class(..content)[length(class(..content))],
        r_class
      )
    )
  ]

  invisible(robopptx)
}

#' Add title to slide candidate
#'
#' @param robopptx robopptx Imported layout file from \link{load_layout}
#' @param title character Title to add to the current slide candidate
#'
#' @return A `robopptx` object.
#' @export
#' @examples "todo"
add_title <- function(robopptx, title = NULL, slide = NA_integer_,
                      shapeid = NA_integer_, hint = NULL) {
  add(
    robopptx = robopptx,
    content = title,
    robo_class = "title",
    slide = slide,
    shapeid = shapeid,
    hint = hint
  )
}

#' Add subtitle to slide candidate
#'
#' @param robopptx robopptx Imported layout file from \link{load_layout}
#' @param sub_title character Subtitle to add to the current slide candidate
#'
#' @return A `robopptx` object.
#' @export
#' @examples "todo"
add_subtitle <- function(robopptx, subtitle = NULL, slide = NA_integer_,
                         shapeid = NA_integer_, hint = NULL) {

  add(
    robopptx = robopptx,
    content = subtitle,
    robo_class = "subtitle",
    slide = slide,
    shapeid = shapeid,
    hint = hint
  )

}

#' Add Text to slide candidate
#'
#' @param robopptx robopptx Imported layout file from \link{load_layout}
#' @param text character Text to add to the current slide candidate
#'
#' @return A `robopptx` object.
#' @export
#' @examples "todo"
add_text <- function(robopptx, text = NULL, slide = NA_integer_,
                     shapeid = NA_integer_, hint = NULL) {
  add(
    robopptx = robopptx,
    content = text,
    robo_class = "text",
    slide = slide,
    shapeid = shapeid,
    hint = hint
  )
}

#' Add Text to slide candidate
#'
#' @param robopptx robopptx Imported layout file from \link{load_layout}
#' @param footer character Footer to add to the current slide candidate
#'
#' @return A `robopptx` object.
#' @export
#' @examples "todo"
add_footer <- function(robopptx, footer = NULL, slide = NA_integer_,
                       shapeid = NA_integer_, hint = NULL) {
  add(
    robopptx = robopptx,
    content = footer,
    robo_class = "footer",
    slide = slide,
    shapeid = shapeid,
    hint = hint
  )
}

#' Add Table to slide candidate
#'
#' @param robopptx robopptx Imported layout file from \link{load_layout}
#' @param table data.frame Table to add to the current slide candidate
#'
#' @return A `robopptx` object.
#' @export
#' @examples "todo"
add_table <- function(robopptx, table = NULL, slide = NA_integer_,
                      shapeid = NA_integer_, hint = NULL) {
  table <- as.data.frame(table)

  add(
    robopptx = robopptx,
    content = table,
    robo_class = "table",
    slide = slide,
    shapeid = shapeid,
    hint = hint
  )
}

#' Add Graph to slide candidate
#'
#' @param robopptx robopptx Imported layout file from \link{load_layout}
#' @param graph ggplot, mschart, charater Graph to add to the current slide candidate as ggplot, mschart or filepath to existing .img or .png file
#'
#' @return A `robopptx` object.
#' @export
#' @examples "todo"
add_graph <- function(robopptx, graph = NULL, slide = NA_integer_,
                      shapeid = NA_integer_, hint = NULL, ...) {
  # if image is given as path - check file and filetype
  if (class(graph)[1] == "character") {
    stopifnot(file.exists(graph))
    stopifnot(grepl("[.](img|png)$", graph))

    add(
      robopptx = robopptx,
      content = graph,
      robo_class = "graph",
      r_class = "external_img",
      slide = slide,
      shapeid = shapeid,
      hint = hint
    )

  } else {
    add(
      robopptx = robopptx,
      content = graph,
      robo_class = "graph",
      slide = slide,
      shapeid = shapeid,
      hint = hint
    )
  }
}

#' Add content to slide candidate
#'
#' @param robopptx robopptx Imported layout file from \link{load_layout}
#' @param ... object Content to add to the current slide candidate
#'
#' @return A `robopptx` object.
#' @export
#' @examples "todo"
add_slide <- function(robopptx, ...) {
  # stopifnot("robopptx" %in% class(robopptx))
  stop_if_not_robopptx(my_layout)


  args <- list(...)

  if (is.null(names(args)) && !is.null(names(unlist(args, recursive = F)))) {
    args <- unlist(args, recursive = F)
  }

  match.arg(names(args), unique(robopptx$robocop$class_mapping$robo_class), several.ok = TRUE)

  robopptx$robocop$add_slide(args)

  invisible(robopptx)
}

#' Add content to slide candidate
#'
#' @param robopptx robopptx Imported layout file from \link{load_layout}
#' @param ... object Content to add to the current slide candidate
#'
#' @return A `robopptx` object.
#' @export
#' @examples "todo"
select_layout <- function(robopptx, ...) {

}
