#' @export
#' @import openxlsx
#' @name copyColumnStyles
#' @param wb A Workbook object containing a worksheet.
#' @param sheet The name of the worksheet to write to.
#' @param colFrom Integer showing the column number in the excel workbook to 
#' copy the formatting from 
#' @param colTo An integer (or vector of integers) defining the column number 
#' in the excel workbook for the format to be copied to
#' @param rows A vector of one or more integer indicating the rows 
#' in the excel workbook to copy the format from
#' @title Copy the formatting of a column to another column for a defined selection of rows

copyColumnStyles <- function(wb, 
                             sheet, 
                             colFrom, 
                             colTo, 
                             rows) {
  
  #wb$styleObjects is a list containing lists containing style, rows and 
  #columns and sheet that style applies to
  #Get styles just from sheet of interest
  all_styles <- wb$styleObjects[sapply(wb$styleObjects, 
                                       function(x) x$sheet == sheet)]
  
  ##Get styles for chosen cols and rows too
  all_styles <- all_styles[sapply(all_styles,
                                  function(x) colFrom %in% x$cols)]
  
  #loop through each style in the workbook
  for(i in 1:length(all_styles)) {
    #For rows in both the style and the specified rows, apply the style
    apply_rows <- intersect(rows, all_styles[[i]]$rows)
    #Then apply the style to that sheet, copying to specified rows and columns
    openxlsx::addStyle(wb,
                       sheet = sheet,
                       style = all_styles[[i]]$style,
                       rows = apply_rows,
                       cols = colTo,
                       stack = TRUE,
                       gridExpand = TRUE)
  }
}
