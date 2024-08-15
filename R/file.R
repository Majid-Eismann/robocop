#' Load PowerPoint to extract its layout
#'
#' @param filepath
#' @param ...
#'
#' @return robopptx
#' @export
#'
#' @examples
load_layout <- function(path = NULL, ...) {
  stopifnot(!is.null(path))
  stopifnot(file.exists(path))

  robopptx <- new_robopptx(path = path, ...)

  return(robopptx)
}

#' Export populated slides to a PowerPoint presentation
#'
#' @param robopptx
#' @param to
#'
#' @return
#' @export
#'
#' @examples
export <- function(robopptx, to = NULL) {
  stopifnot("robopptx" %in% class(robopptx))

  if (length(robopptx$robocop$slides)) {
    for (slide in join_slides(robopptx)) {
      robopptx <- officer::add_slide(robopptx, layout = slide[1, name], master = slide[1, master_name])

      for (element in 1:slide[, .N]) {
        # officer::on_slide(robopptx$layout_pptx, length(robopptx$layout_pptx))

        robopptx <-
          officer::ph_with(
            robopptx,
            value = slide[element, content][[1]][[1]],
            location = officer::ph_location_label(
              type = slide[element, type],
              ph_label = slide[element, ph_label]
            )
          )
      }
    }

    if (is.null(to)) {
      to <- gsub("[.]pptx", "_easypptx.pptx",robopptx$filepath_layout)
    }

    print(robopptx, target = to)
  } else {
    warning("No slides added.")
  }

  return(to)
}
