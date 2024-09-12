#' Group common R classes for a PowerPoint placeholder mapping
#'
#' @param x object
#'
#' @return character robocop content name
#' @keywords internal
#' @export
#' @examples robocontent_dictionary()
robocontent_dictionary <- function() {

  # todo: add typicla r classes

  # positive list of known R classes to be handled by officer and robocop R-package
  dict <-
  list(
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
    tbl = "table"
  )

  dict_table <-
  data.table::data.table(
    r_class = names(dict),
    robo_content = unlist(dict, use.names = F)
  )

  class(dict_table) <- c("robocontentdict", class(dict_table))

  return(dict_table)

}

#' Group common R classes for a PowerPoint placeholder mapping
#'
#' @param x object
#'
#' @return character robocop content name
#' @keywords internal
#' @export
#' @examples "todo"
rclass_to_robo_content <- function(x, abort_if_nomatch = TRUE, return_for_print = FALSE) {

  result <-
  robocontent_dictionary()[
    r_class %in% class(x),
  ]

  if (abort_if_nomatch) stopifnot(result[, uniqueN(robo_content) != 1])

  if (return_for_print) {
    return(result)
  } else return(result[, unique(robo_content)])

}

#' Verify that robocop can draw R-Object on PowerPoint
#'
#' @param x object
#'
#' @return character robocop content name
#' @keywords internal
#' @export
#' @examples can_robocop_handle("My Title")
can_robocop_handle <- function(x) {

  # todo: make pretty cli print
  rclass_to_robo_content(x, abort_if_nomatch = FALSE, return_for_print = TRUE)
}

