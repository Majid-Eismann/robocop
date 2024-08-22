#
# This file contains an improvement for officer::plot_layout_properties() and was submitted as a PR
# for officer: https://github.com/davidgohel/officer/pull/580
#
# FILE CAN BE DELETED ONCE THE PR WAS MERGED INTO OFFICER
#


#' Slide layout properties plot
#'
#' Plot slide layout properties and print informations into defined placeholders.
#' This can be useful to help visualise placeholders locations and identifier.
#'
#' @param x An `rpptx` object.
#' @param layout slide layout name to use
#' @param master master layout name where \code{layout} is located
#' @param labels if \code{TRUE}, placeholder labels will be printed, if \code{FALSE}
#'   placeholder types and identifiers will be printed.
#' @param title if \code{TRUE}, a title with the layout name will be printed.
#' @param shape_id if \code{TRUE}, the shape id of the placeholder will be printed in the upper right
#'   placeholder corner in blue.
#' @importFrom graphics plot rect text box
#' @examples
#' x <- read_pptx()
#' plot_layout_properties(x = x, layout = "Title Slide", master = "Office Theme", title = TRUE, shape_id = TRUE)
#' @family functions for reading presentation informations
#' @export
#' @keywords internal
plot_layout_properties <- function(x, layout = NULL, master = NULL, labels = TRUE, title = TRUE, shape_id = TRUE) {
  old_par <- par(mar = c(2, 2, 1.5, 0))
  on.exit(par(old_par))

  dat <- officer::layout_properties(x, layout = layout, master = master)
  if (length(unique(dat$name)) != 1) {
    stop("one single layout need to be choosen")
  }
  s <- officer::slide_size(x)
  h <- s$height
  w <- s$width
  offx <- dat$offx
  offy <- dat$offy
  cx <- dat$cx
  cy <- dat$cy
  ids <- dat$id
  if (labels) {
    labels <- dat$ph_label
  } else {
    labels <- dat$type[order(as.integer(dat$id))]
    rle_ <- rle(labels)
    labels <- sprintf("type: '%s' - id: %.0f", labels, unlist(lapply(rle_$lengths, seq_len)))
  }
  plot(x = c(0, w), y = -c(0, h), asp = 1, type = "n", axes = FALSE, xlab = NA, ylab = NA)
  if (title) {
    title(main = paste("Layout:", layout))
  }
  rect(xleft = 0, xright = w, ybottom = 0, ytop = -h, border = "darkgrey")
  rect(xleft = offx, xright = offx + cx, ybottom = -offy, ytop = -(offy + cy))
  text(x = offx + cx / 2, y = -(offy + cy / 2), labels = labels, cex = 0.5, col = "red")
  if (shape_id) {
    text(x = offx + cx, y = -offy, labels = ids, cex = 0.6, col = "blue", adj = c(1.2, 1.2), font = 2)
  }
  mtext("y [inch]", side = 2, line = 0, cex = 1.2, col = "darkgrey")
  mtext("x [inch]", side = 1, line = 0, cex = 1.2, col = "darkgrey")
}
