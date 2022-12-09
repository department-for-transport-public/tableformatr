writeDataTable <- function(wb,
                           sheet,
                           x,
                           colNames = TRUE,
                           tableName = NULL,
                           startCol = 1,
                           startRow = 1,
                           keepNA = TRUE,
                           na.string = "[z]",
                           withFilter = FALSE,
                           tableStyle = "none",
                           ...){
  
  ##Just write it normally if keepNA is true
  if(keepNA){
    openxlsx::writeDataTable(wb, 
                            sheet, 
                            x, 
                            startCol = startCol, 
                            startRow = startRow,
                            tableName = tableName,
                            colNames = colNames,
                            keepNA = TRUE,
                            na.string = na.string,
                            withFilter = withFilter,
                            tableStyle = tableStyle,
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
                              withFilter = withFilter,
                              tableStyle = tableStyle,
                              ...)
    
  } else{
    #But if keepNA is a list of columns, we do cool stuff...
    openxlsx::writeData(wb, 
                        sheet, 
                        x, 
                        startCol = startCol, 
                        startRow = startRow,
                        colNames = colNames,
                        ...)
    
    row_values <- 1:nrow(x)
    
    ##Take each bit of the list at a time
    for(i in 1:length(keepNA)){
      for(c in keepNA[[i]]){
        for (r in row_values[is.na(x[,c])]){
          openxlsx::writeData(wb,
                              sheet,
                              rep.NA,
                              startRow = startRow + r - 1,
                              startCol = startCol + c - 1,
                              colNames = FALSE)}
      }
    }
  
  #Get Excel cell references
    ref1 <- paste0(openxlsx:::convert_to_excel_ref(cols = startCol, 
                                                   LETTERS = LETTERS), startRow)
    ref2 <- paste0(openxlsx:::convert_to_excel_ref(cols = startCol + ncol(x) - 1, 
                                                   LETTERS = LETTERS), startRow + nrow(x))
    ref <- paste(ref1, ref2, sep = ":")
  ## replace invalid XML characters
  colNames <- openxlsx:::replaceIllegalCharacters(colNames)
  
  ## create table.xml and assign an id to worksheet tables
  wb$buildTable(
    sheet             = sheet,
    colNames          = colNames,
    ref               = ref,
    showColNames      = showColNames,
    tableStyle        = tableStyle,
    tableName         = tableName,
    withFilter        = withFilter,
    totalsRowCount    = 0L,
    showFirstColumn   = FALSE,
    showLastColumn    = FALSE,
    showRowStripes    = FALSE,
    showColumnStripes = FALSE
  )
  
  }
}