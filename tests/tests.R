rm(list = ls())
library(testthat)
library(robocop)

#### Vignettes
# test_that(
#   "load german_locale.pptx",
#   expect_no_error(
#     load_layout(system.file("extdata", "german_locale.pptx", package = "RoboCop"))
#   )
# )

#### Vignettes
test_that(
  "load sample file",
  expect_no_error({
my_layout <- load_layout(system.file("extdata", "german_locale.pptx", package = "robocop"))

# Add content by calling add_ functions to slide candidate
add_title(my_layout, "Hello RoboCop!")
add_subtitle(my_layout, "slide with add functions")
add_text(my_layout, "RoboCop goes step by step...")

## flush / materialize slide candidate
flush(my_layout)

# Add content by calling add() to slide candidate and let RoboCop guess its content
add(my_layout, "Hello RoboCop!")
add(my_layout, "slide with add() function")
add(my_layout, "RoboCop guessed")

flush(my_layout)

slide(my_layout) <-
  my_layout |>
  add("Hello RoboCop!") |>
  add("slide with adding via piping") |>
  add("RoboCop likes piping")

# # Add content by using magrittr piping and implicitly flushing by assignment
# require(magrittr)
# slide(my_layout) %<>%
# add("Hello RoboCop!") %>%
# add("slide with addding via magrittr piping") %>%
# add("RoboCop likes magrittr")

# Add content by using `+` and implicitly flushing by assignment
slide(my_layout) <-
  my_layout +
  "Hello RoboCop!" +
  "slide with addding via + operator" +
  "RoboCop likes ggplot"


# add a slide at once
add_slide(
  my_layout,
  title = "Hello RoboCop!",
  subtitle = "Slide at once added via add_slide(...)",
  text = "RoboCop functions"
)

# add a slide at once
list(
  title = "Hello RoboCop!",
  subtitle = "Slide at once added via add_slide(...)",
  text = "RoboCop functions"
) |>
  add_slide(robopptx = my_layout)

require(ggplot2)
# add a slide at once
add_slide(
  my_layout,
  title = "ggplot graph",
  graph = ggplot(data = iris, mapping = aes(x = Species, y = Sepal.Length)) +
    geom_point()
)

require(mschart)
add_slide(
  my_layout,
  title = "mschart graph",
  graph = ms_barchart(
    data = browser_data, x = "browser",
    y = "value", group = "serie"
  )
)

add_slide(
  my_layout,
  title = "table",
  table = head(iris)
)

export(my_layout, "RoboCop.pptx")
  })
)
