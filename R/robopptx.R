#' @title Import and prepare PowerPoint file for robocop
#'
#' @description
#' Load layout file, add slides with corresponding contents and export presentation.
#' @return An object of class \code{"robopptx"}.
#' @export
#' @examples "todo"
#' obj <- new_robopptx()
new_robopptx <- function(path = NULL, clean = c("rename", "delete")[1],
                         delete_slides = FALSE, position_precision = 0.5) {
  if (is.null(path)) {
    path <- system.file("extdata", "german_locale.pptx", package = "robocop")
  }

  ## load pptx file via officer
  obj <- officer::read_pptx(path = path)
  obj$filepath_layout <- path

  ## Prepare file
  # clean layout file
  if (clean == "rename") layout_duplicate_rename(obj)
  if (clean == "delete") layout_duplicate_delete(obj)

  # delete sheets
  while (delete_slides && length(obj) > 0) {
    obj$layout_pptx <- officer::remove_slide(obj, 1)
  }

  ## initialize robocop R6 class
  obj$robocop <- robocop$new(rpptx = obj, position_precision = position_precision)

  # add robopptx class
  class(obj) <- c("robopptx", class(obj))

  return(obj)
}


### robocop R6 class -----

#' R6 Class to manage instructions for populating PowerPoint presentation
#'
#' @description
#' Load layout file, add slides with corresponding contents and export presentation.
#' @export
robocop <-
  R6::R6Class(
    "robocop",

    #' @description
    #' Creates a new instance of this [R6][R6::R6Class] class.
    #'
    #' @param rpptx (`rpptx`)\cr
    #'   rpptx object as created by [officer::read_pptx].
    #' @param position_precision (`numeric(1)`)\cr
    #'   TODO
    public = list(
      initialize = function(rpptx, position_precision = 0.5) {
        # stopifnot("rpptx" %in% class(rpptx))
        stop_if_not_rpptx(x)
        ## add class mapping
        # currently supported plot classes
        graph_classes <- c("ggplot", "ms_chart")

        self$class_mapping <-
          list(
            # PowerPoint classes
            class_pptx = c(
              "body", "body", "body",
              "ctrTitle", "subTitle",
              "ftr", "sldNum", "title",
              "pic", "media", "dt"
            ),
            # easy pptx classes
            robo_class = c(
              "text", "table", "graph",
              "title", "subtitle",
              "footer", "slidenumber", "title",
              "graph", "graph", "date"
            ),
            # R classes for mapping
            class_r = c(
              "character", "data.frame", "external_img",
              "character", "character",
              "character", "numeric", "character",
              "external_img", "external_img", "date"
            )
          ) |>
          as.data.table()

        self$class_mapping <-
          rbind(
            self$class_mapping,
            self$class_mapping[
              robo_class == "graph",
            ][
              rep(1:.N, length(graph_classes)),
            ][,
              class_r := rep(graph_classes, each = .N / length(graph_classes))
            ]
          )

        # load layout(s)
        lay_sum <- data.table(officer::layout_summary(rpptx))
        result <- as.list(1:lay_sum[, .N])

        # Looping through each layout
        for (row_index in 1:lay_sum[, .N]) {
          result[[row_index]] <-
            officer::layout_properties(
              rpptx,
              layout = lay_sum[row_index, layout],
              master = lay_sum[row_index, master]
            ) |>
            data.table()

          result[[row_index]][, slide_id := row_index]

          # temporary add slide to get slide size
          officer::add_slide(
            rpptx,
            layout = lay_sum[row_index, layout],
            master = lay_sum[row_index, master]
          )

          # add size information
          size <- officer::slide_size(rpptx)
          result[[row_index]][
            ,
            paste0("layoutslide_", names(size)) := size
          ]

          # remove temporary slide
          officer::remove_slide(rpptx, 1)
        }

        # save layout overview
        self$layout_overview <- rbindlist(result)

        # add robo_classes to map rpptx and R classes
        self$layout_overview <-
          self$layout_overview |>
          merge(
            self$class_mapping,
            by.x = "type",
            by.y = "class_pptx",
            all.x = T,
            allow.cartesian = T
          )

        # mark slides with duplicated powerpoint classes
        self$layout_overview[,
          multi_type := uniqueN(id)>1,
          by = c("slide_id", "type")
        ]

        # divide slide into bottom/top and left/center/right
        self$layout_overview[,
          ":="(
            v_align = ifelse(
              offy >= max(offy) - position_precision,
              "bottom",
              "top"
            ),
            h_align = ifelse(
              # center element
              cx - offx > layoutslide_width * 0.75,
              "center",
              ifelse(
                # width on the right side is bigger then width on the left side
                offx + cx - (layoutslide_width / 2) - position_precision >
                  (layoutslide_width / 2) - position_precision - offx,
                "right",
                "left"
              )
            )
          ),
          by = "slide_id"
        ]

        # Make unique class per slide by combining its layout placeholder classes
        #  and its orientation
        self$layout_overview[,
          slide_class := {
            unique(.SD, by = "id")[, paste0(
              substr(h_align, 1, 1),
              substr(v_align, 1, 1),
              substr(robo_class, 1, 3)
            )
            ] |>
              paste0(collapse = "")
          },
          by = "slide_id"
        ]

        # add layout slide helper columns: number of placeholder, default slide id per class
        self$layout_overview[, placeholder_n := uniqueN(id), by = "slide_id"]
        self$layout_overview[, default_slide_perclass := min(slide_id), by = "slide_class"]

        # slide order
        self$layout_overview[,
          slide_order := match(
            id,
            .SD[
              order(
                factor(v_align, levels = c("top", "bottom")),
                factor(h_align, levels = c("left", "center", "right")),
                offy,
                offx,
                na.last = TRUE
              ),
              id
            ]
          ),
          by = c("slide_id")
        ]

        # save emtpy slide with max number of placeholders
        self$slide_empty <-
          data.table(
            add_order    = NA_integer_,
            robo_class   = NA_character_,
            class_r      = NA_character_,
            content      = list(list()),
            dotdotdot    = list(list()),
            hint         = list(list()),
            user_selection_slide = NA_integer_,
            user_selection_shapeid = NA_integer_
          )[rep(1, self$layout_overview[, max(placeholder_n)])]

        # empty slide in current slide slot
        self$slide_candidate <- copy(self$slide_empty)
      },

      #' @field layout_pptx (`TODO`)\cr
      #' TODO
      layout_pptx = NULL,

      #' @field layout_overview (`data.frame`)\cr
      #' TODO
      layout_overview = data.table(NULL),

      #' @field class_mapping (`data.frame`)\cr
      #' TODO
      class_mapping = data.table(NULL),

      #' @description TODO
      #' @param layout (`TODO`)\cr TODO
      show_layout = function(layout) {

      },

      #' @description TODO
      #' @param slide (`TODO`)\cr TODO
      show_slide = function(slide) {

      },

      #' @description TODO
      show_slide_candidate = function() {

      },

      #' @field slide_candidate (`data.frame`)\cr
      #' TODO
      slide_candidate = data.table(NULL),

      #' @field slide_empty (`data.frame`)\cr
      #' TODO
      slide_empty = data.table(NULL),

      #' @field slides (`list`)\cr
      #' TODO
      slides = list(),

      #' @description TODO
      flush_candidate = function() {
        # add slide to slidedeck
        self[["slides"]][[length(self$slides) + 1]] <- self$slide_candidate[!is.na(add_order)]

        # add emtpy instruction set
        self$slide_candidate <- copy(self$slide_empty)
      },

      #' @description TODO
      join_slides = function() {

        # self <- my_layout$robocop
        # .x <- 2

        if (length(self$slides)) {
          map(
            1:length(self$slides),
            ~ {

              # map slide and layout by robo_class and class_r by layout slide_id
              self$layout_overview[
                placeholder_n >= self$slides[[.x]][, .N],
                list(id, slide_class, placeholder_n, slide_order, default_slide_perclass, class_count = order(slide_order)),
                by = c("robo_class", "slide_id", "class_r")
              ] |>
                merge(
                  self$slides[[.x]][
                    !is.na(robo_class),
                    list(slide_filter = T, class_count = order(add_order)),
                    by = c("class_r", "robo_class")
                  ],
                  by = c("robo_class", "class_count", "class_r"),
                  all.x = T
                ) ->
              layout_selection

              # complete robo_class(es)
              if (self$slides[[.x]][, any(is.na(robo_class))]) {
                layout_selection <-
                  layout_selection |>
                  merge(
                    self$slides[[.x]][is.na(robo_class), list(add_order, class_r, slide_filter = T)],
                    by.x = c("class_r", "slide_order"),
                    by.y = c("class_r", "add_order"),
                    all.x = T
                  )

                layout_selection[,
                  slide_filter := slide_filter.x | slide_filter.y
                ]
                layout_selection[, ":="(slide_filter.x = NULL, slide_filter.y = NULL)]
              }

              # complete layout selection
              selected_layout <-
                layout_selection[,
                  sum(slide_filter, na.rm = T),
                  by = c("slide_id", "slide_class", "default_slide_perclass", "placeholder_n")
                ][V1 == max(V1) & default_slide_perclass == slide_id, ][placeholder_n == min(placeholder_n), ][slide_id == min(slide_id), slide_id]

              self$slides[[.x]][, list(add_order, class_r, content)] %>%
                merge(
                  self$layout_overview[
                    slide_id == selected_layout,
                  ],
                  by.x = c("add_order", "class_r"),
                  by.y = c("slide_order", "class_r")
                )
            }
          )
        } else {
          warnings("No slides to join. Forgot to add content and/or flush slide(s)?")
        }
      },

      #' @description TODO
      #' @param content_list (`list`)\cr
      #'  TODO
      add_slide = function(content_list = list()) {

        if (length(unlist(content_list))) {
          # laod empty slide
          tmp <- copy(self$slide_empty)

          # add content
          tmp[
            1:length(content_list),
            ":="(
              add_order = 1:length(content_list),
              robo_class = names(content_list),
              class_r = map_chr(content_list, ~class(.x)[length(class(.x))]),
              content = setNames(content_list, NULL)
            )
          ]

          # save slide
          self[["slides"]][[length(self$slides) + 1]] <- tmp[!is.na(add_order)]
        } else {
          warning("Empty list passed to add_slide method from easy_pptx")
        }
      },

      #' @description TODO
      #'  TODO
      explain = function() {

      }
    )
  )
