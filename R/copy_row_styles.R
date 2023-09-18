#' @export
#' @import openxlsx
#' @name copyRowStyles
#' @param wb A Workbook object containing a worksheet.
#' @param sheet The name of the worksheet to write to.
#' @param rowFrom Integer showing the row number in the excel workbook to 
#' copy the formatting from 
#' @param rowTo An integer (or vector of integers) defining the row number 
#' in the excel workbook for the format to be copied to
#' @param cols A vector of one or more integer indicating the cols 
#' in the excel workbook to copy the format from
#' @title Copy the formatting of a row to another row for a defined selection of columns

copyRowStyles <- function(wb, 
                             sheet, 
                             rowFrom, 
                             rowTo, 
                             cols) {
  
  #wb$styleObjects is a list containing lists containing style, rows and 
  #columns and sheet that style applies to
  #Get styles just from sheet of interest
  all_styles <- wb$styleObjects[sapply(wb$styleObjects, 
                                       function(x) x$sheet == sheet)]
  
  ##Get styles for chosen cols and rows too
  all_styles <- all_styles[sapply(all_styles,
                                  function(x) rowFrom %in% x$rows)]
  
  #loop through each style in the workbook
  for(i in 1:length(all_styles)) {
    #For rows in both the style and the specified rows, apply the style
    apply_cols <- intersect(cols, all_styles[[i]]$cols)
    #Then apply the style to that sheet, copying to specified rows and columns
    openxlsx::addStyle(wb,
                       sheet = sheet,
                       style = all_styles[[i]]$style,
                       rows = rowTo,
                       cols = apply_cols,
                       stack = TRUE,
                       gridExpand = TRUE)
  }
}
