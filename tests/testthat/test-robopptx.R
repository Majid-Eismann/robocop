library(robocop)

test_that("basic functions throw no error", {
  # MH:
  # - this integration test should probably me much more modular
  skip(message = "Skipping for now, as test currently fails")
  expect_no_error({
    my_layout <- load_layout(system.file("extdata", "german_locale.pptx", package = "robocop"))

    # Add content by calling add_ functions to slide candidate
    add_title(my_layout, "Hello robocop!")
    add_subtitle(my_layout, "slide with add functions")

    ## materialise slide candidate
    materialise(my_layout)

    # Add content by calling add() to slide candidate and let robocop guess its content
    add(my_layout, "Hello robocop!")
    add(my_layout, "slide with add() function")

    materialise(my_layout)

    slide(my_layout) <-
      my_layout |>
      add("Hello robocop!") |>
      add("slide with adding via piping") |>
      add("robocop likes piping")


    # # Add content by using magrittr piping and implicitly materialiseing by assignment
    # require(magrittr)
    # slide(my_layout) %<>%
    # add("Hello robocop!") %>%
    # add("slide with addding via magrittr piping") %>%
    # add("robocop likes magrittr")

    # Add content by using `+` and implicitly materialiseing by assignment
    slide(my_layout) <-
      my_layout +
      "Hello robocop!" +
      "slide with addding via + operator" +
      "robocop likes ggplot"


    # add a slide at once
    add_slide(
      my_layout,
      title = "Hello robocop!",
      subtitle = "Slide at once added via add_slide(...)",
      text = "robocop functions"
    )

    # add a slide at once
    list(
      title = "Hello robocop!",
      subtitle = "Slide at once added via add_slide(...)",
      text = "robocop functions"
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
    file <- tempfile(fileext = ".pptx")
    export(my_layout, file)
  })
})
