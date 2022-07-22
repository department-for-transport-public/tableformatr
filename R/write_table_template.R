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
#' @title Write to a worksheet template and format as an Excel table

writeDataTableTemplate <- function(wb,
                                   sheet,
                                   x,
                                   startRow,
                                   tableName = NULL,
                                   startCol = 1,
                                   colNames = TRUE,
                                   rowNames = FALSE,
                                   keepNA = openxlsx_getOp("keepNA", FALSE),
                                   na.string = openxlsx_getOp("na.string")
                                   ){

  ##Read in first row of data to become the headers
  names <- openxlsx::readWorkbook(wb,
                                  rows =  startRow,
                                  colNames = FALSE)

  ##If there aren't enough titles for rows, error out
  if(length(names) != ncol(x)){
    stop(paste("Incorrect number of column names provided in template.",
               ncol(x), "names are needed,",
               length(names),
               "are provided in template"))
  }

  #Rename table with titles from template
  names(x) <- names[1,]


  ##Write data out
  openxlsx::writeDataTable(wb,
                           sheet = sheet,
                           x = x,
                           tableStyle = "none",
                           startRow = startRow,
                           tableName = tableName,
                           startCol = startCol,
                           colNames = colNames,
                           rowNames = rowNames,
                           keepNA = keepNA,
                           na.string = na.string
  )

}

