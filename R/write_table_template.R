#' @export
#' @import openxlsx
#' @name writeDataTableTemplate
#' @param wb A Workbook object containing a worksheet.
#' @param sheet The worksheet to write to. Can be the worksheet index or name.
#' @param x A dataframe
#' @param startRow A vector specifying the starting row to write df
#' @param tableName name of table in workbook. The table name must be unique.
#' @param startCol A vector specifying the starting column to write df
#' @param colNames If TRUE, column names of x are written.
#' @param rowNames If TRUE, row names of x are written.
#' @param keepNA If TRUE, NA values are converted to #N/A (or na.string, if not NULL) in Excel, else NA cells will be empty.
#' @param na.string If not NULL, and if keepNA is TRUE, NA values are converted to this string in Excel.
#' @param withFilter whether to include filters on the table or not. Defaults to FALSE
#' @title Write to a worksheet template and format as an Excel table

writeDataTableTemplate <- function(wb,
                                   sheet,
                                   x,
                                   startRow,
                                   tableName = NULL,
                                   startCol = 1,
                                   colNames = TRUE,
                                   rowNames = FALSE,
                                   keepNA = TRUE,
                                   na.string = "[z]",
                                   withFilter = FALSE
                                   ){

  ##Read in first row of data to become the headers
  names <- openxlsx::readWorkbook(wb,
                                  sheet = sheet,
                                  rows =  startRow:(startRow+1),
                                  skipEmptyRows = FALSE,
                                  skipEmptyCols = FALSE,
                                  colNames = FALSE)
  
  ##Start names from start column
  names <- names[1, startCol:length(names)]

  ##If there aren't enough titles for rows, error out
  if(length(names) < ncol(x)){
    stop(paste("Incorrect number of column names provided in template.",
               ncol(x), "names are needed,",
               length(names),
               "are provided in template"))
  } 
  
  ##If there are too many names, only keep the ones we need but warn
  if(length(names) > ncol(x)){
    
    names <- names[1:ncol(x)]
    
    
  }

  #Rename table with titles from template
  names(x) <- unlist(names)


  ##Write data out
  writeReplaceDataTable(wb,
                           sheet = sheet,
                           x = x,
                           startRow = startRow,
                           tableName = tableName,
                           startCol = startCol,
                           colNames = colNames,
                           rowNames = rowNames,
                           keepNA = keepNA,
                           na.string = na.string,
                          withFilter = withFilter
  )

}

