#' @export
#' @name superscriptNotes
#' @param wb A Workbook object containing one or more worksheets.
#' @title Convert footnote annotations using the format [note x] to superscript
#' @import openxlsx

superscriptNotes <- function(wb){

  for(i in grep("\\[((N|n)ote.*)\\]", wb$sharedStrings)){

    # insert additional formatting in shared string
    wb$sharedStrings[[i]] <- gsub("<si>",
                                  "<si><r>",
                                  gsub("</si>",
                                       "</r></si>",
                                       wb$sharedStrings[[i]]))

    # find the "[...]" pattern, remove brackets and underline and enclose the text with superscript format
    wb$sharedStrings[[i]] <- gsub("\\[((N|n)ote.*)\\]",
                                  "</t></r><r><rPr><vertAlign val=\"superscript\"/></rPr><t xml:space=\"preserve\">\\[\\1\\]</t></r><r><t xml:space=\"preserve\">",
                                  wb$sharedStrings[[i]])
   }
}

