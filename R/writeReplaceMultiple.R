#' @export
#' @import openxlsx
#' @name writeReplaceMultiple
#' @param wb A Workbook object containing a worksheet.
#' @param sheet The worksheet to write to. Can be the worksheet index or name.
#' @param x A dataframe
#' @param colNames If TRUE, column names of x are written.
#' @param tableName name of table in workbook. The table name must be unique.
#' @param startRow A vector specifying the starting row to write df
#' @param startCol A vector specifying the starting column to write df
#' @param keepNA If TRUE, NA values are converted to #N/A 
#' (or na.string, if not NULL) in Excel. If FALSE, NA cells will be empty. 
#' @param na.string If keepNA is TRUE, all NA values are converted to this string in Excel.
#' @param keepZero If FALSE, 0 values are converted to the zero.string. If TRUE, they will remain as 0.
#' @param zero.string If keepZero is FALSE, all zero values are converted to this string in Excel
#' @param withFilter whether to include filters on the table or not. Defaults to FALSE
#' @param ... Additional arguments to be passed to writeDataTable
#' @title Write to an Excel table replacing NA values in a custom way

writeReplaceMultiple <- function(wb,
                                  sheet,
                                  x,
                                  colNames = TRUE,
                                  tableName = NULL,
                                  startCol = 1,
                                  startRow = 1,
                                  keepNA = TRUE,
                                  na.string = "[z]",
                                  keepZero = TRUE,
                                  zero.string = NULL,
                                  withFilter = FALSE,
                                  ...){
  ##Remove data.table formatting that breaks this
  x <- as(x, "data.frame")
  
  ##Write the data out
  openxlsx::writeDataTable(wb, 
                           sheet, 
                           x, 
                           startCol = startCol, 
                           startRow = startRow,
                           tableName = tableName,
                           colNames = colNames,
                           withFilter = FALSE,
                           tableStyle = "none",
                           ...)
  if(keepNA == TRUE) {

      row_values <- 1:nrow(x)
      
      
      ##Take each bit of the list at a time
        ##Replace your NA values
        for(c in 1:ncol(x)){
          for (r in row_values[is.na(x[, c])]){
            openxlsx::writeData(wb,
                                sheet,
                                na.string,
                                startRow = startRow + r,
                                startCol = startCol + c - 1,
                                colNames = FALSE)}
        }
      
      
  }
  
  ##Ditto for zero values
  if(keepZero == FALSE) {
    
    #If keepZero is false, we replace stuff
    row_values <- 1:nrow(x)
    
    
    ##Take each bit of the list at a time
    ##Replace your NA values
    for(c in 1:ncol(x)){
      
      ##Calculate row values, and if they equal NA, drop them
      row_values <- row_values[!is.na(x[, c])]
      
      if(length(row_values) != 0){
      
        for (r in row_values[x[, c] == 0]){
          openxlsx::writeData(wb,
                              sheet,
                              zero.string,
                              startRow = startRow + r,
                              startCol = startCol + c - 1,
                              colNames = FALSE)}
      }
     
    } 
    
  }
}