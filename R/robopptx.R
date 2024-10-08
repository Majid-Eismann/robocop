### robopptx  -----


#' @title Import and prepare PowerPoint file for robocop
#'
#' @description
#' Load layout file, add slides with corresponding contents and export presentation.
#' @return An object of class \code{"robopptx"}.
#' @export
#' @examples "todo"
#' obj <- new_robopptx()
#' @rawNamespace import(purrr, except = transpose)
new_robopptx <- function(path = NULL, clean = "rename",
                         delete_slides = FALSE, position_precision = 0.5) {
  if (is.null(path)) {
    path <- system.file("extdata", "german_locale.pptx", package = "robocop")
  }
  clean <- match.arg(clean, c("rename", "delete"))
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

#' Get mapping table for PowerPoint types, robocop content and r classes
#'
#' @param supported_graph_classes object
#'
#' @return data.table data.table holding class mapping
#' @keywords internal
#' @export
#' @examples "todo"
get_mapping_table <- function(supported_graph_classes = c("ggplot", "ms_chart")) {
  # hard coded mapping table
  mapping_table <-
    list(
      # PowerPoint classes
      class_pptx = c(
        "body", "body", "body",
        "ctrTitle", "subTitle",
        "ftr", "sldNum", "title",
        "pic", "media", "dt"
      ),
      # robocop content
      robo_content = c(
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

  # recycle robocop content for certain r classes
  rbind(
    mapping_table,
    mapping_table[
      robo_content == "graph",
    ][
      rep(1:.N, length(supported_graph_classes)),
    ][
      ,
      class_r := rep(supported_graph_classes, each = .N / length(supported_graph_classes))
    ]
  )
}

#' Add alignment to layout table
#'
#' @param layout_table data.table describing layouts returned by \link{get_layout_table}
#'
#' @return data.table
#' @keywords internal
#' @examples
#' "todo"
add_alignment_layout_table <- function(layout_table, position_precision = 0.5) {
  # todo: footer v_align

  # divide slide into bottom/top and left/center/right
  layout_table[,
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
    by = "layout_id"
  ]
}

#' Add Layout Combination of Contents to layout Table
#'
#' @param layout_table data.table describing layouts returned by \link{get_layout_table}
#'
#' @return character name of grouped class
#' @keywords internal
#' @examples
#' "todo"
add_layoutcombination_layout_table <- function(layout_table) {
  # Make unique class per slide by combining its layout placeholder classes
  #  and its orientation
  layout_table[,
    layout_class := {
      unique(.SD, by = "id")[, paste0(
        substr(h_align, 1, 1),
        substr(v_align, 1, 1),
        substr(robo_content, 1, 3)
      )] |>
        paste0(collapse = "")
    },
    by = "layout_id"
  ]
}


#' Add Order of every placement per Layout
#'
#' @param layout_table data.table describing layouts returned by \link{get_layout_table}
#'
#' @return character name of grouped class
#' @keywords internal
#' @examples
#' add_placementorder_layout_table(get_layout_table())
add_placementorder_layout_table <- function(layout_table) {
  # slide order
  layout_table[,
    placement_order := match(
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
    by = c("layout_id")
  ]
}

#' Add necessary variables for robocop to layer table
#'
#' @param layout_table data.table describing layouts returned by \link{get_layout_table}
#'
#' @return character name of grouped class
#' @keywords internal
#' @examples
#' add_robocop_variables(get_layout_table())
add_robocop_variables <- function(layout_table, position_precision = 0.5) {
  # expand layout PowerPoint types by every robocop content and r-class
  layout_table <-
    merge(
      layout_table,
      get_mapping_table(),
      by.x = "type",
      by.y = "class_pptx",
      all.x = T,
      allow.cartesian = T
    )

  # add vertical and horizontal alignment to placeholder by layout
  # (top,bottom,footer and left,center,right)
  add_alignment_layout_table(layout_table, position_precision = position_precision)

  # add layout combination code
  add_layoutcombination_layout_table(layout_table)

  # number of placeholder
  layout_table[, placeholder_n := uniqueN(id), by = "layout_id"]

  # default slide id per class
  layout_table[, default_slide_perclass := min(layout_id), by = "layout_class"]

  # add order of placements
  add_placementorder_layout_table(layout_table)
}

#' Get mapping table for PowerPoint types, robocop content and r classes
#'
#' @param x object
#'
#' @return character name of grouped class
#' @keywords internal
#' @export
#' @examples "todo"
get_layout_table <- function(rpptx, position_precision = 0.5) {
  # get all layout placeholders from officer
  layout_table <-
    rpptx$slideLayouts$get_xfrm_data() |>
    merge(
      rpptx$slideLayouts$get_metadata() |>
        subset(select = -c(master_file)),
      by = c("name", "master_name")
    ) |>
    data.table()

  # add layout_id from numeration of layout files
  layout_table[
    ,
    layout_id := as.numeric(stringr::str_extract(filename, "\\d+"))
  ]

  # check layout id uniqueness &
  stopifnot(all(layout_table[, uniqueN(layout_id), by = "filename"]$V1 == 1))
  stopifnot(layout_table[, all(!is.na(layout_id))])

  # add slide size to every layout placeholder
  slide_size <- officer::slide_size(rpptx)
  layout_table[
    ,
    paste0("layoutslide_", names(slide_size)) := slide_size
  ]

  # add additional variables needed by robocop to layout
  layout_table <-
    add_robocop_variables(layout_table, position_precision = position_precision)
}

#' Get mapping table for PowerPoint types, robocop content and r classes
#'
#' @param number_of_rows number of rows with default values
#'
#' @return data.table
#' @keywords internal
#' @export
#' @examples "todo"
get_emptyslide_table <- function(number_of_rows = 1) {
  data.table(
    add_order = NA_integer_,
    robo_content = NA_character_,
    class_r = NA_character_,
    content = list(list()),
    dotdotdot = list(list()),
    hint = list(list()),
    user_selection_layout = NA_integer_,
    user_selection_shapeid = NA_character_
  )[rep(1, number_of_rows)]
}

#' @export
#' @keywords internal
print.robopptx <- function(x, ...) {
  cli::cli_h2("robopptx")
  df_layouts <- officer::layout_summary(x)
  masters <- df_layouts$master |> unique()
  layouts <- df_layouts$layout |> unique()
  n_layouts <- df_layouts |> nrow()
  n_slides <- x$slide$length()
  bullets <- c(
    "masters: {.val {length(masters)}} [{.val {masters}}]",
    "layouts: {.val {n_layouts}} [{.val {layouts}}]",
    "slides: {.val {n_slides}}"
  )
  names(bullets) <- rep("*", length(bullets))
  cli::cli_bullets(bullets)
}

#' @export
#' @keywords internal
print.robocontentdict <- function(x, ...) {
  cli::cli_h2("robocop dictionary of contents")

}


#' Materialise current slide candidate into a slide
#'
#' @param robopptx robopptx Imported layout file from [load_layout]
#'
#' @return A `robopptx` object.
#' @export
#' @examples "todo"
materialise <- function(robopptx) {
  # stopifnot("robopptx" %in% class(robopptx))
  stop_if_not_robopptx(robopptx)

  # materialise candidate
  robopptx$robocop$materialise_candidate()

  invisible(robopptx)
}

#' Materialise all slides into a presentation
#'
#' @param robopptx robopptx Imported layout file from [load_layout]
#'
#' @return TODO
#' @export
#' @examples "todo"
materialise_presentation <- function(robopptx) {
  # stopifnot("robopptx" %in% class(robopptx))
  stop_if_not_robopptx(robopptx)
  robopptx$robocop$materialise_presentation()
}

#' Internal slide function to select possible layout(s)
#'
#' @param slide single slide to select layout for
#' @param robocop_r6 R6
#'
#' @return TODO
#' @examples "todo"
slide_layout_candidate <- function(slide, robocop_r6) {

  stop_if_not_class(robocop_r6, "R6")

  # if user selected layout
  if (slide[, any(!is.na(user_selection_layout))]) {

    stopifnot(slide[!is.na(user_selection_layout), uniqueN(user_selection_layout) == 1])

    layout_candidate <-
      robocop_r6$layout_overview[
        layout_id == slide[!is.na(user_selection_layout), user_selection_layout[1]],
        list(id, layout_class, placeholder_n, placement_order, default_slide_perclass, class_count = order(placement_order)),
        by = c("robo_content", "layout_id", "class_r")
      ]
  } else {

    layout_candidate <-
      robocop_r6$layout_overview[
        placeholder_n >= slide[, .N],
        list(id, layout_class, placeholder_n, placement_order, default_slide_perclass, class_count = order(placement_order)),
        by = c("robo_content", "layout_id", "class_r")
      ]
  }

  return(list(result = layout_candidate, decision = list()))
}

#' Internal slide function to match users content and placeholders
#'
#' @param slide single slide to select layout for
#' @param layout_candidate
#' @param robocop_r6 R6
#'
#' @return TODO
#' @examples "todo"
slide_match_placeholder <- function(slide, layout_candidate, robocop_r6) {

  # match slide and layout via robo_content and class_r by layout slide_id
  #   robo_content is given that robocop does not have to guess
  layout_candidate |>
    merge(
      slide[
        !is.na(robo_content) & is.na(user_selection_shapeid),
        list(shape_weight = 1, class_count = order(add_order)),
        by = c("class_r", "robo_content")
      ],
      by = c("robo_content", "class_count", "class_r"),
      all.x = T
    ) ->
    matched_slide

  # if user selected ids
  if (slide[, any(!is.na(user_selection_shapeid))]) {
    # append user selected ids (without matching)
    matched_slide <-
      rbind(
        matched_slide,
        layout_candidate |>
          merge(
            slide[
              !is.na(user_selection_shapeid),
              list(shape_weight = 100),
              by = list(id = user_selection_shapeid)
            ],
            by = "id",
            all.y = T
          )
      )
  }

  # robo content is not given - robocop has to guess what content it is
  if (slide[, any(is.na(robo_content))]) {
    matched_slide <-
      matched_slide |>
      merge(
        slide[is.na(robo_content), list(add_order, class_r, shape_weight = 1)],
        by.x = c("class_r", "placement_order"),
        by.y = c("class_r", "add_order"),
        all.x = T
      )

    matched_slide[
      ,
      shape_weight := sum(shape_weight.x, shape_weight.y) - 1
    ]
    matched_slide[, ":="(shape_weight.x = NULL, shape_weight.y = NULL)]
  }

  return(list(result = matched_slide, decision = list()))

}

#' Internal slide function to select final layout
#'
#' @param slide single slide to select layout for
#' @param matched_slide
#'
#' @return TODO
#' @examples "todo"
slide_pick_layout <- function(slide, matched_slide) {

  # if user has picked the layout
  if (slide[, any(!is.na(user_selection_layout))]) {
    selected_layout <-
      slide[!is.na(user_selection_layout), user_selection_layout[1]]
  # otherwise let robocop pick a layout
  } else {
    #
    selected_layout <-
      matched_slide[,
        # sum the number of matched placeholders (1 for each matched, 100 for a user selected shapeid)
        sum(shape_weight, na.rm = T),
        by = c("layout_id", "layout_class", "default_slide_perclass", "placeholder_n")
      ][
        # select layout with most matched placeholders and use its default class
        V1 == max(V1) & default_slide_perclass == layout_id,
      ][layout_id == min(layout_id), layout_id]
  }

  return(list(result = selected_layout, decision = list()))

}

#' Internal slide function to select final layout
#'
#' @param slide
#' @param robocop_r6 robopptx Imported layout file from [load_layout]
#' @param selected_layout
#'
#' @return TODO
#' @examples "todo"
slide_finalize <- function(slide, robocop_r6, selected_layout) {

  rbind(
    ##!!! TODO has to be matched_placeholder filtered for selected_layout
    slide[is.na(user_selection_shapeid), list(add_order, class_r, content)] |>
      merge(
        robocop_r6$layout_overview[
          layout_id == selected_layout,
        ],
        by.x = c("add_order", "class_r"),
        by.y = c("placement_order", "class_r")
      ),
    slide[!is.na(user_selection_shapeid), list(content, id = user_selection_shapeid)] |>
      merge(
        robocop_r6$layout_overview[
          layout_id == selected_layout,
        ],
        by = "id"
      ) |>
      data.table::setnames("placement_order", "add_order")
  )

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
        stop_if_not_rpptx(rpptx)

        # get mapping table: map PowerPoint types, robocop content and r classes
        self$class_mapping <- get_mapping_table()

        # load layout(s)
        self$layout_overview <- get_layout_table(rpptx, position_precision = position_precision)

        # save emtpy slide with max number of placeholders
        self$slide_empty <- get_emptyslide_table(self$layout_overview[, max(placeholder_n)])

        # empty slide in current slide slot
        self$slide_candidate <- copy(self$slide_empty)
      },

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

      #' @field cursor (`list`)\cr
      #' TODO
      add_constraint = list(),

      #' @description TODO
      get_layout_overview = function() {
        # if any constraints
        if (length(self$add_constraint)) {
          # filter layout_overview
          self$layout_overview[
            add_constraint,
            on = names(add_constraint)
          ]
        } else {
          self$layout_overview
        }
      },

      #' @description TODO
      materialise_candidate = function() {
        # add slide to slidedeck
        self[["slides"]][[length(self$slides) + 1]] <- self$slide_candidate[!is.na(add_order)]

        # replace all layout ids with the last added entry
        if (self[["slides"]][[length(self$slides)]][, any(!is.na(user_selection_layoutid))])
        self[["slides"]][[length(self$slides)]][
          ,
          user_selection_layoutid := .SD[!is.na(user_selection_layoutid), ][add_order == max(add_order), user_selection_layoutid]
        ]

        # reset add constraint
        self$add_constraint <- list()

        # reset to an empty instruction set
        self$slide_candidate <- copy(self$slide_empty)
      },

      #' @description TODO
      materialise_presentation = function() {
        # self <- my_layout$robocop
        # .x <- 2

        # if any slides
        if (length(self$slides)) {

            purrr::map(
              self$slides,
              ~{

                # refactor and make it plotable/explainable
                #> possible_layouts
                # 1. Select possible layouts / User Layout <Performance
                #    - if all robo_content defined and layout matches perfectly, take the layout
                layout_candidate <-
                  slide_layout_candidate(slide = .x, robocop_r6 = self)

                #> match shapes
                # 2. Merge slide and all layouts (by robo_content, r_class) <
                # 3. Append user ids
                # 4. robocop guess elements
                matched_slide <-
                  slide_match_placeholder(
                    slide = .x,
                    robocop_r6 = self,
                    layout_candidate = layout_candidate$result
                  )

                #> decide layout
                # 5. Decide which layout to take
                picked_layout <-
                  slide_pick_layout(slide = .x, matched_slide = matched_slide$result)

                #> finalize_slide
                # 6. Add content to the layout
                slide_finalize(
                  slide = .x,
                  robocop_r6 = self,
                  selected_layout = picked_layout$result
                )

              }
            )

        } else {
          warnings("No slides to join. Forgot to add content and/or materialise slide(s)?")
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
              robo_content = names(content_list),
              class_r = purrr::map_chr(content_list, ~ class(.x)[length(class(.x))]),
              content = stats::setNames(content_list, NULL)
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
