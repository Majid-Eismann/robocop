#' Group common R classes for a PowerPoint placeholder mapping
#'
#' @param x object
#'
#' @return character robocop content name
#' @keywords internal
#' @import purrr
#' @export
#' @examples "todo"
rclass_to_robo_content <- function(x) {

  result <-
    purrr::map_chr(
      class(x),
      switch,
      character = "text",
      POSIXct = "date",
      POSIXt = "date",
      Date = "date",
      ## graph
      ggplot = "graph",
      gg = "graph",
      ms_chart = "graph",
      ## table
      data.frame = "table",
      data.table = "table",
      tbl_df = "table",
      tbl= "table",
      NA_character_
    ) |>
    unique() |>
    na.omit()


  stopifnot(length(result) == 1)

  return(result)
}

#' Flush / materialize current slide candidate into a slide
#'
#' @param robopptx robopptx Imported layout file from [load_layout]
#'
#' @return A `robopptx` object.
#' @export
#' @examples "todo"
flush <- function(robopptx) {
  # stopifnot("robopptx" %in% class(robopptx))
  stop_if_not_robopptx(robopptx)

  # flush candidate
  robopptx$robocop$flush_candidate()

  invisible(robopptx)
}

#' Flush / materialize current slide candidate into a slide
#'
#' @param robopptx robopptx Imported layout file from [load_layout]
#'
#' @return TODO
#' @export
#' @examples "todo"
join_slides <- function(robopptx) {
  # stopifnot("robopptx" %in% class(robopptx))
  stop_if_not_robopptx(robopptx)
  robopptx$robocop$join_slides()
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
  flush(x)
  x
})
