# Function to save a plot to PDF. 
# plot_function must be a defined plotting function &
# filename will be the designated file name for the pdf
save.pdf <- function(plot_function, filename) {
  # Define the output PDF file path
  pdf_file <- file.path("Output", paste0(filename, ".pdf"))
  # Open the PDF device to save the plot
  pdf(pdf_file)
  # Call the plot function to create the plot
  print(plot_function())
  # Close the PDF device (this finalizes and writes the plot to the file)
  dev.off()
  # Inform the user that the plot has been saved
  message("Plot has been saved to: ", pdf_file)
}

# =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
plot_range <- function(data, value_col, y_label, plot_title) {
  ggplot(data, aes(x = Time.Category, y = .data[[value_col]])) +
    geom_boxplot(fill = "skyblue", alpha = 0.7) +
    labs(
      title = plot_title,
      x = "Time of Day",
      y = y_label
    ) +
    theme_minimal()
}


# =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
# Old functions
plot_range2 <- function(data, max_col, min_col, y_label, plot_title) {
  ggplot(data, aes(x = Time.of.Day, group = Day, color = Day)) +
    geom_line(aes(y = .data[[max_col]]), size = 1) +
    geom_point(aes(y = .data[[max_col]])) +
    geom_line(aes(y = .data[[min_col]]), linetype = "dashed", size = 1) +
    geom_point(aes(y = .data[[min_col]]), shape = 1) +
    labs(
      title = plot_title,
      x = "Time of Day",
      y = y_label,
      color = "Day"
    ) +
    scale_x_time(
      limits = c(hms::as_hms("07:30:00"), hms::as_hms("16:30:00")),
      breaks = hms::as_hms(c("08:00:00", "12:00:00", "16:00:00"))
    ) +
    theme_minimal()
}