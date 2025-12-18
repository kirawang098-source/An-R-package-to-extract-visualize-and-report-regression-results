

# R/reg_extract.R

#' Extract regression coefficients
#'
#' @description
#' Extracts coefficients, standard errors, t/z values, and p-values from
#' an lm or glm object and returns a tidy data.frame.
#'
#' @param model An lm or glm object
#'
#' @return A data.frame with columns: term, estimate, std.error, statistic, p.value
#'
#' @examples
#' lm_model <- lm(mpg ~ wt + hp, data = mtcars)
#' reg_extract(lm_model)
#'
#' @export
reg_extract <- function(model) {
  if (!requireNamespace("broom", quietly = TRUE)) {
    stop("Please install the 'broom' package to use this function.")
  }
  broom::tidy(model)
}

#' Plot regression coefficients
#'
#' @description
#' Creates a forest-style plot of regression coefficients with 95% CI.
#'
#' @param model An lm or glm object
#' @param show_p Logical, whether to highlight significant coefficients (p < 0.05)
#'
#' @return A ggplot object
#'
#' @examples
#' lm_model <- lm(mpg ~ wt + hp, data = mtcars)
#' reg_plot(lm_model)
#'
#' @export
reg_plot <- function(model, show_p = TRUE) {
  if (!requireNamespace("broom", quietly = TRUE)) {
    stop("Please install the 'broom' package to use this function.")
  }
  if (!requireNamespace("ggplot2", quietly = TRUE)) {
    stop("Please install the 'ggplot2' package to use this function.")
  }

  df <- broom::tidy(model)
  df$signif <- ifelse(df$p.value < 0.05, "Yes", "No")

  p <- ggplot2::ggplot(df, ggplot2::aes(x = estimate, y = term)) +
    ggplot2::geom_point(ggplot2::aes(color = signif), size = 3) +
    ggplot2::geom_errorbarh(
      ggplot2::aes(xmin = estimate - 1.96 * std.error,
                   xmax = estimate + 1.96 * std.error),
      height = 0.2
    ) +
    ggplot2::scale_color_manual(values = c("Yes" = "red", "No" = "black")) +
    ggplot2::labs(x = "Estimate", y = "Term", color = "Significant") +
    ggplot2::theme_minimal()

  return(p)
}

#' Generate Word report of regression results
#'
#' @description
#' Creates a Word document containing a regression coefficient table and a forest plot.
#'
#' @param model An lm or glm object
#' @param file_path Output file path, e.g., "result.docx"
#'
#' @return A Word document saved to file_path
#'
#' @examples
#' lm_model <- lm(mpg ~ wt + hp, data = mtcars)
#' reg_report(lm_model, "reg_result.docx")
#' reg_report(lm_model, "reg_result.docx")
#'
#' @export
reg_report <- function(model, file_path = "reg_result.docx") {
  if (!requireNamespace("broom", quietly = TRUE) ||
      !requireNamespace("flextable", quietly = TRUE) ||
      !requireNamespace("officer", quietly = TRUE) ||
      !requireNamespace("ggplot2", quietly = TRUE)) {
    stop("Please install 'broom', 'flextable', 'officer', and 'ggplot2'.")
  }

  # coefficients
  df <- broom::tidy(model)

  # tables
  ft <- flextable::qflextable(df)

  # generate the graph
  df$signif <- ifelse(df$p.value < 0.05, "Yes", "No")
  p <- ggplot2::ggplot(df, ggplot2::aes(x = estimate, y = term)) +
    ggplot2::geom_point(ggplot2::aes(color = signif), size = 3) +
    ggplot2::geom_errorbarh(
      ggplot2::aes(xmin = estimate - 1.96 * std.error,
                   xmax = estimate + 1.96 * std.error),
      height = 0.2
    ) +
    ggplot2::scale_color_manual(values = c("Yes" = "red", "No" = "black")) +
    ggplot2::labs(x = "Estimate", y = "Term", color = "Significant") +
    ggplot2::theme_minimal()

  # 新建 Word 文档
  doc <- officer::read_docx()
  doc <- officer::body_add_par(doc, "Regression Results", style = "heading 1")
  doc <- flextable::body_add_flextable(doc, ft)
  doc <- officer::body_add_par(doc, "Coefficient Plot", style = "heading 1")
  doc <- officer::body_add_gg(doc, value = p, width = 6, height = 4)

  print(doc, target = file_path)

  message("Word report saved to: ", file_path)
}

