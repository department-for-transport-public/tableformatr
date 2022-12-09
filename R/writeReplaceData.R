#' @export
#' @import openxlsx
#' @name writeReplaceDataTable
#' @param wb A Workbook object containing a worksheet.
#' @param sheet The worksheet to write to. Can be the worksheet index or name.
#' @param x A dataframe
#' @param colNames If TRUE, column names of x are written.
#' @param tableName name of table in workbook. The table name must be unique.
#' @param startRow A vector specifying the starting row to write df
#' @param startCol A vector specifying the starting column to write df
#' @param keepNA If TRUE, NA values are converted to #N/A 
#' (or na.string, if not NULL) in Excel. If FALSE, NA cells will be empty. 
#' A list of column numbers can also be passed to this argument; in this case
#' NA values in the first columns in the list will be converted to the 
#' first string provided to na.string, the second to the second string, etc.
#' @param na.string If keepNA is TRUE, all NA values are converted to this string in Excel.
#' If keepNA is a list of column numbers a vector of values can be provided to this argument, 
#' NA values in the specified columns are converted to their corresponding string 
#' @title Write to an Excel table replacing NA values in a custom way

writeReplaceDataTable <- function(wb,
                           sheet,
                           x,
                           colNames = TRUE,
                           tableName = NULL,
                           startCol = 1,
                           startRow = 1,
                           keepNA = TRUE,
                           na.string = "[z]",
                           ...){
  if(!is.list(keepNA)){
    ##Just write it normally if keepNA is true
    if(keepNA == TRUE){
      openxlsx::writeDataTable(wb, 
                              sheet, 
                              x, 
                              startCol = startCol, 
                              startRow = startRow,
                              tableName = tableName,
                              colNames = colNames,
                              keepNA = TRUE,
                              na.string = na.string,
                              withFilter = FALSE,
                              tableStyle = "none",
                              ...)
    } else if(keepNA == FALSE){
      ##Just don't replace NA if its false
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
      
    }} else{
      
      ##Check if the list of columns and NAs is the same
      if(length(keepNA) != length(na.string)){
        stop("Length of keepNA list and na.string vector must be equal")
      }
      
    #But if keepNA is a list of columns, we do cool stuff...
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
    
    row_values <- 1:nrow(x)
  
    
    ##Take each bit of the list at a time
    for(i in 1:length(keepNA)){
      ##Replace your NA values
      for(c in keepNA[[i]]){
        for (r in row_values[is.na(x[,c])]){
          openxlsx::writeData(wb,
                              sheet,
                              na.string[[i]],
                              startRow = startRow + r,
                              startCol = startCol + c - 1,
                              colNames = FALSE)}
      }
    }
  
  
  }
}