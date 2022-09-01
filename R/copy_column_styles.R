#' @export
#' @import openxlsx
#' @name copyColumnStyles
#' @param wb A Workbook object containing a worksheet.
#' @param sheet The worksheet to write to. Can be the worksheet index or name.
#' @param colFrom Integer showing the column number in the excel workbook of which the format is to be copied
#' @param colTo An Integer defining the column number in the excel workbook for the format to be copied to
#' @param startRow An integer indicating the row in the excel workbook to start copying the format
#' @param endRow An integer indicating the row in the excel workbook to stop copying the format
#' @title Copy the formatting of a column to another column for a defined consecutive selection of rows

copyColumnStyles <- function(wb, sheet, colFrom, colTo, startRow, endRow) {
  
  #wb$styleObjects is a list containing lists containing style, rows and columns and sheet that style applies to
  all_styles <- wb$styleObjects
  
  #loop through each row
  for(row in startRow:endRow) {
    #loop through each style
    for(i in 1:length(all_styles)) {
      #if that style applies to the row looking at and the column copying from
      if(all_styles[[i]]$sheet==sheet & row %in% all_styles[[i]]$rows & colFrom %in% all_styles[[i]]$cols) {
        #Then apply the style to that row and column copying to
        openxlsx::addStyle(wb,sheet, all_styles[[i]]$style, row,colTo,stack = TRUE)
      }
    }
  }
}
