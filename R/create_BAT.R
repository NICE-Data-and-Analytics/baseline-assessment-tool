
library(dplyr)
library(openxlsx)
library(officer)

# guidance_content <- read_docx("NG226 Guideline 20221019.docx") %>% 
#         scrape_docx()
# 
# guidance_content <- readWorkbook(
#                         guidance_content$wb,
#                         sheet = 1,
#                         startRow = 1,
#                         colNames = TRUE,
#                         rowNames = FALSE,
#                         skipEmptyRows = TRUE,
#                         skipEmptyCols = TRUE)

scrape_title_and_dates <- function(doc){
    
    temp_df <- docx_summary(doc)
    
    temp <- temp_df %>% 
        filter(style_name %in% c("Title 1", "Guidance issue date", "Document issue date")) %>% 
        select(text) %>% 
        pull() %>% 
        str_trim()
    
    return(temp)
}

# Set up cell styles ---------------------------------------------------------

intro_title_style <- createStyle(
    fontName = "Lora Semibold", fontSize = 22, fontColour = "#000000",
    halign = "Left", valign = "top", wrapText = TRUE)

intro_date_style <- createStyle(
    fontName = "Inter", fontSize = 18, fontColour = "#000000", 
    halign = "Left", valign = "top")

datasheet_title_style <- createStyle(
    fontName = "Lora Semibold", fontSize = 22, fontColour = "#000000",
    halign = "Left", valign = "top")

header_style <- createStyle(
    fontName = "Inter Semibold", fontSize = 13, fontColour = "#FFFFFF", 
    halign = "Left", valign = "top", fgFill = "#00436C")

subheader_style <- createStyle(
    fontName = "Inter Semibold", fontSize = 12, fontColour = "#FFFFFF", 
    halign = "Left", valign = "top", fgFill = "#228096")

text_style <- createStyle(
    fontName = "Inter", fontSize = 12, fontColour = "#000000", 
    halign = "Left", valign = "top", wrapText = TRUE)

highlight_style <- createStyle(
    fontName = "Inter", fontSize = 12, fontColour = "#000000", 
    halign = "Left", valign = "top", fgFill = "#EAD054", wrapText = TRUE)

border_style <- createStyle(
    border = c("top", "bottom", "left", "right"),
    borderStyle = c("thin", "thin", "thin", "thin"),
    borderColour = "#000000")

hyperlink_style <- createStyle(
    fontName = "Inter", fontSize = 12, fontColour = "#005EA5", 
    halign = "Left", valign = "top", textDecoration = "underline")


# Set up functions needed to adjust row heights ---------------------------

# The code below was taken from the following open pull request to the openxlsx package
# https://github.com/awalker89/openxlsx/pull/382/files

#' @name get_worksheet_entries
#' @title Get entries from workbook worksheet
#' @description Get all entries from workbook worksheet without xml tags
#' @param wb workbook
#' @param sheet worksheet
#' @author David Breuer
#' @return vector of strings
#' @export
#' @examples
#' ## Create new workbook
#' wb <- createWorkbook()
#' addWorksheet(wb, "Sheet")
#' sheet <- 1
#'
#' ## Write dummy data
#' writeData(wb, sheet, c("A", "BB", "CCC"), startCol = 2, startRow = 3)
#' writeData(wb, sheet, c(4, 5), startCol = 4, startRow = 3)
#'
#' ## Get text entries
#' get_worksheet_text(wb, sheet)
#'
get_worksheet_entries <- function(wb, sheet) {
    # get worksheet data
    dat <- wb$worksheets[[sheet]]$sheet_data
    # get vector of entries
    val <- dat$v
    # get boolean vector of text entries
    typ <- (dat$t == 1) & !is.na(dat$t)
    # get text entry strings
    str <- unlist(wb$sharedStrings[as.integer(val)[typ] + 1])
    # remove xml tags
    str <- gsub("<.*?>", "", str)
    # write strings to vector of entries
    val[typ] <- str
    # return vector of entries
    val
}

#' @name auto_heights
#' @title Compute optimal row heights
#' @description Compute optimal row heights for cell with fixed with and
#' enabled automatic row heights parameter
#' @param wb workbook
#' @param sheet worksheet
#' @param selected selected rows
#' @param fontsize font size, optional (get base font size by default)
#' @param factor factor to manually adjust font width, e.g., for bold fonts,
#' optional
#' @param base_height basic row height, optional
#' @param extra_height additional row height per new line of text, optional
#' @author David Breuer
#' @return list of indices of columns with fixed widths and optimal row heights
#' @export
#' @examples
#' ## Create new workbook
#' wb <- createWorkbook()
#' addWorksheet(wb, "Sheet")
#' sheet <- 1
#'
#' ## Write dummy data
#' long_string <- "ABC ABC ABC ABC ABC ABC ABC ABC ABC ABC ABC"
#' writeData(wb, sheet, c("A", long_string, "CCC"), startCol = 2, startRow = 3)
#' writeData(wb, sheet, c(4, 5), startCol = 4, startRow = 3)
#'
#' ## Set column widths and get optimal row heights
#' setColWidths(wb, sheet, c(1,2,3,4), c(10,20,10,20))
#' auto_heights(wb, sheet, 1:5)
#'
auto_heights <- function(wb, sheet, selected, fontsize = NULL, factor = 1.0,
                         base_height = 15, extra_height = 12) { #base from 15, extra from 12
    # get base font size
    if (is.null(fontsize)) {
        fontsize <- as.integer(openxlsx::getBaseFont(wb)$size$val)
    }
    # set factor to adjust font width (empiricially found scale factor 4 here)
    factor <- 4 * factor / fontsize
    # get worksheet data
    dat <- wb$worksheets[[sheet]]$sheet_data
    # get columns widths
    colWidths <- wb$colWidths[[sheet]]
    # select fixed (non-auto) and visible (non-hidden) columns only
    specified <- (colWidths != "auto") & (attr(colWidths, "hidden") == "0")
    # return default row heights if no column widths are fixed
    if (length(specified) == 0) {
        message("No column widths specified, returning default row heights.")
        cols <- integer(0)
        heights <- rep(base_height, length(selected))
        return(list(cols, heights))
    }
    # get fixed column indices
    cols <- as.integer(names(specified)[specified])
    # get fixed column widths
    widths <- as.numeric(colWidths[specified])
    # get all worksheet entries
    val <- get_worksheet_entries(wb, sheet)
    # compute optimal height per selected row
    heights <- sapply(selected, function(row) {
        # select entries in given row and columns of fixed widths
        index <- (dat$rows == row) & (dat$cols %in% cols)
        # remove line break characters
        chr <- gsub("\\r|\\n", "", val[index])
        # measure width of entry (in pixels)
        wdt <- strwidth(chr, unit = "in") * 20 / 1.43 # 20 px = 1.43 in
        # compute optimal height
        if (length(wdt) == 0) {
            base_height
        } else {
            base_height + extra_height * as.integer(max(wdt / widths * factor))
        }
    })
    # return list of indices of columns with fixed widths and optimal row heights
    list(cols, heights)
}

#'
#' ## Write dummy data
#' writeData(wb, sheet, "fixed w/fixed h", startCol = 1, startRow = 1)
#' writeData(wb, sheet, "fixed w/auto h ABC ABC ABC ABC ABC ABC ABC ABC ABC ABC ABC", startCol = 2, startRow = 2)
#' writeData(wb, sheet, "variable w/fixed h", startCol = 3, startRow = 3)
#' 
#' ## Set column widths and row heights
#' setColWidths(wb, sheet, cols = c(1, 2, 3, 4), widths = c(10, 20, "auto", 20))
#' setRowHeights(wb, sheet, rows = c(1, 2, 8, 4, 6), heights = c(30, "auto", 15, 15, 30))
#' 
#' ## Overwrite row 1 height
#' setRowHeights(wb, sheet, rows = 1, heights = 40)
#' 
#' ## Save workbook
#' saveWorkbook(wb, "setRowHeightsExample.xlsx", overwrite = TRUE)
#' 
setRowHeights <- function(wb, sheet, rows, heights,
                          fontsize = NULL, factor = 1.0,
                          base_height = 15, extra_height = 12, wrap = TRUE) {
    # validate sheet
    sheet <- wb$validateSheet(sheet)
    if (length(rows) > length(heights))
        heights <- rep(heights, length.out = length(rows))
    if (length(heights) > length(rows))
        stop("Greater number of height values than rows.")
    od <- getOption("OutDec")
    options(OutDec = ".")
    on.exit(expr = options(OutDec = od), add = TRUE)
    # clean duplicates
    heights <- heights[!duplicated(rows)]
    rows <- rows[!duplicated(rows)]
    # auto adjust row heights
    ida <- which(heights == "auto")
    selected <- rows[ida]
    out <- auto_heights(wb, sheet, selected, fontsize = fontsize, factor = factor,
                        base_height = base_height, extra_height = extra_height)
    cols <- out[[1]]
    new <- out[[2]]
    heights[ida] <- new
    names(heights) <- rows
    # wrap text in cells
    if (wrap == TRUE) {
        wrap <- openxlsx::createStyle(wrapText = TRUE)
        openxlsx::addStyle(wb, sheet, wrap, rows = ida, cols = cols, gridExpand = T, stack = T)
    }
    wb$setRowHeights(sheet, rows, heights)
}


# Create BAT --------------------------------------------------------------

create_BAT <- function(guidance_number, guidance_info, guidance_content){
    
    
    guidance_content <- readWorkbook(
                                guidance_content,
                                sheet = 1,
                                startRow = 1,
                                colNames = TRUE,
                                rowNames = FALSE,
                                skipEmptyRows = TRUE,
                                skipEmptyCols = TRUE)
    ### Create all of the variables to drop into the BAT ###
    # NOTE: Everything has to be in tibble or data frame format #
    
    intro_title <- paste0("Baseline assessment tool for ", guidance_info[[1]], " (", guidance_number, ")")
    intro_published_date <- paste0("Published: ", guidance_info[[2]])
    
    if (length(guidance_info) == 3) {
        intro_update_date <- paste0("Updated: ", guidance_info[[3]])
    }
    
    # Formula to create a hyperlink to the guidance
    guidance_hyperlink <- tibble(
        link = paste0(
            "HYPERLINK(\"",
            "https://www.nice.org.uk/guidance/", guidance_number,
            "\", \"", guidance_info[[1]], "\")"))
    
    # Formula to create a hyperlink the the tools and resources tab
    tools_hyperlink <- tibble(
        link = paste0(
            "HYPERLINK(\"",
            "https://www.nice.org.uk/guidance/", guidance_number, "/resources",
            "\", \"Tools and resources\")"))
    
    # Formula to create a hyperlink the the Notice of Rights
    rights_hyperlink <- tibble(
        link = paste0(
            "HYPERLINK(\"",
            "https://www.nice.org.uk/terms-and-conditions#notice-of-rights",
            "\", \"Subject to Notice of rights\")"))
    
    # Set up the formulas to drop into the 'Data totals' tab
    # These are set up to adjust the formula to the correct number of recommendations
    datatotal_formulas <- tibble(
        formula = c(paste0("=SUMPRODUCT(COUNTIF('Data sheet'!C3:C", nrow(guidance_content)+2, ",{\"Yes\",\"Partial\"}))"),
                    paste0("=COUNTIF('Data sheet'!E3:E", nrow(guidance_content)+2, ",\"Yes\")"),
                    paste0("=COUNTIF('Data sheet'!E3:E", nrow(guidance_content)+2, ",\"Partial\")")))
    
    # These need to be converted to a class of formula so that excel recognizes them
    class(guidance_hyperlink$link) <- "formula"
    class(tools_hyperlink$link) <- "formula"
    class(rights_hyperlink$link) <- "formula"
    class(datatotal_formulas$formula) <- "formula"
    
    ### Load the template and drop in all of the relevant content ###
    
    wb <- loadWorkbook("BAT template.xlsx")
    
    writeData(wb, sheet = "Introduction", intro_title, startRow = 1, startCol = 1, colNames = FALSE)
    writeData(wb, sheet = "Introduction", intro_published_date, startRow = 2, startCol = 1, colNames = FALSE)
    writeData(wb, sheet = "Introduction", guidance_hyperlink, startRow = 5, startCol = 1, colNames = FALSE)
    writeData(wb, sheet = "Introduction", tools_hyperlink, startRow = 10, startCol = 1, colNames = FALSE)
    writeData(wb, sheet = "Introduction", rights_hyperlink, startRow = 12, startCol = 1, colNames = FALSE)
    writeData(wb, sheet = "Data sheet", intro_title, startRow = 1, startCol = 1, colNames = FALSE)
    writeData(wb, sheet = "Data sheet", guidance_content, startRow = 3, startCol = 1, colNames = FALSE)
    writeData(wb, sheet = "Data sheet totals", datatotal_formulas, startRow = 1, startCol = 2, colNames = FALSE)
    deleteData(wb, sheet = "Data sheet", cols = 2, rows = str_which(guidance_content$rec_number, "Heading|Subheading|Text")+2, gridExpand = TRUE)
    
    if (length(guidance_info) == 3) {
        writeData(wb, sheet = "Introduction", intro_update_date, startRow = 3, startCol = 1, colNames = FALSE)
        addStyle(wb, sheet = "Introduction", intro_date_style, rows = 3, cols = 1, stack = FALSE)
    }
    
    ### Apply all of the appropriate styles ###
    
    addStyle(wb, sheet = "Introduction", intro_title_style, rows = 1, cols = 1, stack = FALSE)
    addStyle(wb, sheet = "Introduction", intro_date_style, rows = 2, cols = 1, stack = FALSE)
    addStyle(wb, sheet = "Introduction", hyperlink_style, rows = 5, cols = 1, stack = FALSE)
    addStyle(wb, sheet = "Introduction", hyperlink_style, rows = 10, cols = 1, stack = TRUE)
    addStyle(wb, sheet = "Introduction", hyperlink_style, rows = 12, cols = 1, stack = FALSE)
    addStyle(wb, sheet = "Data sheet", datasheet_title_style, rows = 1, cols = 1, stack = FALSE)
    addStyle(wb, sheet = "Data sheet", header_style, rows = str_which(guidance_content$rec_number, "Heading")+2, cols = 1:12, stack = FALSE, gridExpand = TRUE)
    addStyle(wb, sheet = "Data sheet", subheader_style, rows = str_which(guidance_content$rec_number, "Subheading")+2, cols = 1:12, stack = FALSE, gridExpand = TRUE)
    addStyle(wb, sheet = "Data sheet", text_style, rows = str_which(guidance_content$rec_number, "[:digit:]")+2, cols = 1:12, stack = FALSE, gridExpand = TRUE)
    addStyle(wb, sheet = "Data sheet", highlight_style, rows = which(guidance_content$rec_number == "Text")+2, cols = 1:12, stack = FALSE, gridExpand = TRUE)
    addStyle(wb, sheet = "Data sheet", border_style, rows = str_which(guidance_content$rec_number, "[:digit:]|Text")+2, cols = 1:12, stack = TRUE, gridExpand = TRUE)
    
    ### Add in the drop downs and conditional formatting ###
    
    #setRowHeights(wb, sheet = 2, rows = str_which(guidance_content$rec_number, "[:digit:]|Text")+2, heights = "auto", wrap = FALSE, base_height = 20, extra_height = 11)
    dataValidation(wb, sheet = "Data sheet", rows = str_which(guidance_content$rec_number, "[:digit:]")+2, cols = c(3,5), type = "list", value = "'Dropdowns'!$A$1:$A$3")
    dataValidation(wb, sheet = "Data sheet", rows = str_which(guidance_content$rec_number, "[:digit:]")+2, cols = 7, type = "list", value = "'Dropdowns'!$A$1:$A$2")
    conditionalFormatting(wb, sheet = "Data sheet", rows = str_which(guidance_content$rec_number, "[:digit:]")+2, cols = 4:11, rule = '$C4="No"', style = createStyle(bgFill = "#BFBFBF"))
    
    ### Save the complete BAT ###
    
    return(wb)
    #saveWorkbook(wb,"output.xlsx", overwrite = TRUE)
}

# create_BAT(guidance_number = "NG226",
#            guidance_title = "Osteoarthritis in over 16s: diagnosis and management",
#            published_date = "19 October 2022",
#            guidance_content = guidance_content)
