#' Add content to layout candidate
#'
#' @param robopptx robopptx Imported layout file from \link{load_layout}
#' @param content object Content to add to the current layout candidate
#'
#' @return A `robopptx` object.
#' @export
#'
#' @examples "todo"
add <- function(robopptx, content = NULL, layoutid = NA_integer_, shapeid = NA_character_,
                hint = NULL, ..., robo_content = NA_character_, r_class = NA_character_) {
  stop_if_not_robopptx(robopptx)
  stop_if_not_correct_shapeid(shapeid, robopptx, robo_content = robo_content)
  stop_if_not_correct_layoutid(layoutid, robopptx, robo_content = robo_content)

  dots <- list(...)

  add_position <- robopptx$robocop$slide_candidate[!is.na(add_order), .N] + 1

  # set cursor layout id
  if (!is.na(layoutid)) {
    robopptx$robocop$add_constraint$layout_id <- as.integer(layoutid)
  } else if (!is.null(robopptx$robocop$add_constraint$layout_id)) {
    layoutid <- robopptx$robocop$add_constraint$layout_id
  }

  robopptx$robocop$slide_candidate[
    add_position,
    ":="(
      add_order = add_position,
      content = list(list(..content)),
      user_selection_layoutid = as.integer(layoutid),
      user_selection_shapeid = as.character(shapeid),
      hint = list(list(hint)),
      dotdotdot = list(list(dots)),
      robo_content = ifelse(
        is.na(..robo_content),
        NA_character_,
        ..robo_content
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

#' Add title to layout candidate
#'
#' @param robopptx robopptx Imported layout file from \link{load_layout}
#' @param title character Title to add to the current layout candidate
#'
#' @return A `robopptx` object.
#' @export
#' @examples "todo"
add_title <- function(robopptx, title = NULL, layoutid = NA_integer_,
                      shapeid = NA_integer_, hint = NULL) {
  add(
    robopptx = robopptx,
    content = title,
    robo_content = "title",
    layoutid = layoutid,
    shapeid = shapeid,
    hint = hint
  )
}

#' Add subtitle to layout candidate
#'
#' @param robopptx robopptx Imported layout file from \link{load_layout}
#' @param sub_title character Subtitle to add to the current layout candidate
#'
#' @return A `robopptx` object.
#' @export
#' @examples "todo"
add_subtitle <- function(robopptx, subtitle = NULL, layoutid = NA_integer_,
                         shapeid = NA_integer_, hint = NULL) {
  add(
    robopptx = robopptx,
    content = subtitle,
    robo_content = "subtitle",
    layoutid = layoutid,
    shapeid = shapeid,
    hint = hint
  )
}

#' Add Text to layout candidate
#'
#' @param robopptx robopptx Imported layout file from \link{load_layout}
#' @param text character Text to add to the current layout candidate
#'
#' @return A `robopptx` object.
#' @export
#' @examples "todo"
add_text <- function(robopptx, text = NULL, layoutid = NA_integer_,
                     shapeid = NA_integer_, hint = NULL) {
  add(
    robopptx = robopptx,
    content = text,
    robo_content = "text",
    layoutid = layoutid,
    shapeid = shapeid,
    hint = hint
  )
}

#' Add Text to layout candidate
#'
#' @param robopptx robopptx Imported layout file from \link{load_layout}
#' @param footer character Footer to add to the current layout candidate
#'
#' @return A `robopptx` object.
#' @export
#' @examples "todo"
add_footer <- function(robopptx, footer = NULL, layoutid = NA_integer_,
                       shapeid = NA_integer_, hint = NULL) {
  add(
    robopptx = robopptx,
    content = footer,
    robo_content = "footer",
    layoutid = layoutid,
    shapeid = shapeid,
    hint = hint
  )
}

#' Add Table to layout candidate
#'
#' @param robopptx robopptx Imported layout file from \link{load_layout}
#' @param table data.frame Table to add to the current layout candidate
#'
#' @return A `robopptx` object.
#' @export
#' @examples "todo"
add_table <- function(robopptx, table = NULL, layoutid = NA_integer_,
                      shapeid = NA_integer_, hint = NULL) {
  table <- as.data.frame(table)

  add(
    robopptx = robopptx,
    content = table,
    robo_content = "table",
    layoutid = layoutid,
    shapeid = shapeid,
    hint = hint
  )
}

#' Add Graph to layout candidate
#'
#' @param robopptx robopptx Imported layout file from \link{load_layout}
#' @param graph ggplot, mschart, charater Graph to add to the current layout candidate as ggplot, mschart or filepath to existing .img or .png file
#'
#' @return A `robopptx` object.
#' @export
#' @examples "todo"
add_graph <- function(robopptx, graph = NULL, layoutid = NA_integer_,
                      shapeid = NA_integer_, hint = NULL, ...) {
  # if image is given as path - check file and filetype
  if (class(graph)[1] == "character") {
    stopifnot(file.exists(graph))
    stopifnot(grepl("[.](img|png)$", graph))

    add(
      robopptx = robopptx,
      content = graph,
      robo_content = "graph",
      r_class = "external_img",
      layoutid = layoutid,
      shapeid = shapeid,
      hint = hint
    )
  } else {
    add(
      robopptx = robopptx,
      content = graph,
      robo_content = "graph",
      layoutid = layoutid,
      shapeid = shapeid,
      hint = hint
    )
  }
}

#' Add content to layout candidate
#'
#' @param robopptx robopptx Imported layout file from \link{load_layout}
#' @param ... object Content to add to the current layout candidate
#'
#' @return A `robopptx` object.
#' @export
#' @examples "todo"
add_slide <- function(robopptx, ...) {
  stop_if_not_robopptx(my_layout)

  args <- list(...)

  if (is.null(names(args)) && !is.null(names(unlist(args, recursive = F)))) {
    args <- unlist(args, recursive = F)
  }

  match.arg(names(args), unique(robopptx$robocop$class_mapping$robo_content), several.ok = TRUE)

  robopptx$robocop$add_slide(args)

  invisible(robopptx)
}

## Add methods for base operators

#' Adding content with `+` operator
#'
#' @param e1 robopptx Imported layout file from [load_layout]
#' @param e2 object Content to add with [add]
#'
#' @return A `robopptx` object.
#' @keywords internal
#' @export
"+.robopptx" <- function(e1, e2) add(robopptx = e1, content = e2)

#' Register an R6 class as an S4 class
#'
#' This makes the R6 class "robopptx" available as an S4 class.
#'
#' # @export
#'
#' # setOlxdClass(c("robopptx")

#' Method for the '+' generic for objects of class 'robopptx'
#'
#' @param e1 An object of class 'robopptx'
#' @param e2 object Content to add with [add]
#'
#' @return A `robopptx` object.
#' @keywords internal
#' @export
setMethod(
  "+",
  signature = c("robopptx", "ANY"),
  definition = function(e1, e2) RoboCop::add(robopptx = e1, content = e2)
)

# `<-` operator


#' @export
setGeneric("slide", function(x) standardGeneric("slide"))

#' Method for the 'myFunction' generic for objects of class 'MyClass'
#'
#' @param robopptx robopptx Imported layout file from [load_layout]
#' @return A `robopptx` object.
#' @export
setMethod("slide", "robopptx", function(x) x$slide)

#' @export
setGeneric("slide<-", function(x, value) standardGeneric("slide<-"))

#' @export
setMethod("slide<-", "robopptx", function(x) {
  materialise(x)
  x
})

#' Add content to layout candidate
#'
#' @param robopptx robopptx Imported layout file from \link{load_layout}
#' @param ... object Content to add to the current layout candidate
#'
#' @return A `robopptx` object.
#' @export
#' @examples "todo"
select_layout <- function(robopptx, layoutid) {
  stop_if_not_correct_layoutid(layoutid, robopptx)

  robopptx$robocop$add_constraint$layout_id <- as.integer(layoutid)

  invisible(robopptx)
}
