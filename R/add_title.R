#' @export
#' @name addTitle
#' @param wb A Workbook object.
#' @param title A string containing the title for the Workbook
#' @param overwrite boolean specifying if the title should overwrite any existing title. Defaults to FALSE
#' @title Add a title to a workbook core information, which will display in the file information tab in Excel.


addTitle <- function(wb, title, overwrite = FALSE){

  #Check to see if there is a title already
  if(grepl("<dc:title>", wb$core, fixed = TRUE)){
    if(overwrite == FALSE){
      stop("Workbook already has a title. Choose overwrite=TRUE to overwrite existing title")
    }else if(overwrite == TRUE){
      #Add core tags round title
      title_string <- paste0("<dc:title>",
                             title,
                             "</dc:title>")
      ##Gsub out the creator with a title
      wb$core <- gsub("[<]dc\\:title\\>.*\\<*dc\\:title\\>",
                      title_string,
                      wb$core)
    }

    }else{
  #Add core tags round title
  title_string <- paste0("<dc:title>",
                         title,
                         "</dc:title><dc:creator>")
  ##Gsub out the creator with a title
  wb$core <- gsub("<dc:creator>",
                  title_string,
                  wb$core,
                  fixed = TRUE)
  }
}
