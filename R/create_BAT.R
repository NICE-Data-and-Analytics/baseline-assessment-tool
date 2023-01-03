
library(dplyr)
library(openxlsx)
library(officer)


# Testing section ---------------------------------------------------------

# doc <- read_docx("files/test_files/NG228 Guideline_for prep.docx")
# 
# guidance_number <- "NG288"
# guidance_info <- scrape_title_and_dates(doc)
# guidance_content <- scrape_docx(doc)
# 
# guidance_content <- tibble(
#     readWorkbook(
#         guidance_content$wb,
#         sheet = 1,
#         startRow = 1,
#         colNames = TRUE,
#         rowNames = FALSE,
#         skipEmptyRows = TRUE,
#         skipEmptyCols = TRUE))
# 
# saveWorkbook(test,"output.xlsx", overwrite = TRUE)


# Create function to gather guideline title and dates ---------------------

scrape_title_and_dates <- function(doc){
    
    temp_df <- docx_summary(doc)
    
    # Title can be styled as "Title" or "Title 1"
    temp <- temp_df %>% 
        filter(style_name %in% c("Title", 
                                 "Title 1", 
                                 "Guidance issue date", 
                                 "Document issue date")) %>% 
        select(text) %>% 
        pull() %>% 
        stringr::str_trim()
    
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


# Create BAT --------------------------------------------------------------

create_BAT <- function(guidance_number, guidance_info, guidance_content){
    
    # Read in the pre-scraped content from the guideline
    guidance_content <- tibble(
                            readWorkbook(
                                guidance_content,
                                sheet = 1,
                                startRow = 1,
                                colNames = TRUE,
                                rowNames = FALSE,
                                skipEmptyRows = TRUE,
                                skipEmptyCols = TRUE))
    
    # Sort out the year of recommendation column. Whenever there is an NA next to a rec number 
    # insert the original year of publication for the guideline
    guidance_content <- guidance_content %>% 
        filter(!rec_number %in% c("Panel", "Caption")) %>% 
        mutate(rec_year = ifelse((is.na(rec_year) & !rec_number %in% c("Heading", "Subheading",
                                                                       "Subsubheading", "Text")),
                                 str_sub(guidance_info[2], -4,-1),
                                 rec_year))
    
    ### Create all of the variables to drop into the BAT ###
    # NOTE: These are in individuals tibbles because they would not format correctly 
    # when not in tibble format. They would also lose their formatting when all links
    # were in a single tibble - no idea why though.
    
    intro_title <- paste0("Baseline assessment tool for ", 
                          str_to_lower(str_sub(guidance_info[1], 1, 1)),
                          str_sub(guidance_info[1], 2),
                          " (", guidance_number, ")")
    
    intro_published_date <- paste0("Published: ", guidance_info[2])
    
    if (length(guidance_info) == 3) {
        intro_update_date <- paste0("Updated: ", guidance_info[3])
    }
    
    # Formula to create a hyperlink to the guidance
    guidance_hyperlink <- tibble(
        link = paste0('HYPERLINK(\"',
                      'https://www.nice.org.uk/guidance/', guidance_number,
                      '\", \"', guidance_info[1], '\")'))
    
    # Formula to create a hyperlink the the tools and resources tab
    tools_hyperlink <- tibble(
        link = paste0('HYPERLINK(\"',
                      'https://www.nice.org.uk/guidance/', guidance_number, '/resources',
                      '\", \"Tools and resources\")'))
    
    # Formula to create a hyperlink the the Notice of Rights
    rights_hyperlink <- tibble(
        link = paste0('HYPERLINK(\"',
                      'https://www.nice.org.uk/terms-and-conditions#notice-of-rights',
                      '\", \"Subject to Notice of rights\")'))
    
    # Set up the formulas to drop into the 'Data totals' tab
    # These are set up to adjust the formula to the correct number of recommendations
    datatotal_formulas <- tibble(
        formula = c(paste0("=SUMPRODUCT(COUNTIF('Data sheet'!D3:D", 
                           nrow(guidance_content)+2, ',{\"Yes\",\"Partial\"}))'),
                    paste0("=COUNTIF('Data sheet'!F3:F", nrow(guidance_content)+2, ',\"Yes\")'),
                    paste0("=COUNTIF('Data sheet'!F3:F", nrow(guidance_content)+2, ',\"Partial\")')))
    
    # These need to be converted to a class of formula so that excel recognizes them
    class(guidance_hyperlink$link) <- "formula"
    class(tools_hyperlink$link) <- "formula"
    class(rights_hyperlink$link) <- "formula"
    class(datatotal_formulas$formula) <- "formula"
    
    
    ### Load the template and drop in all of the relevant content ###
    
    wb <- loadWorkbook("files/input_files/BAT_template.xlsx")
    
    writeData(wb, sheet = "Introduction", intro_title, 
              startRow = 1, startCol = 1, colNames = FALSE)
    
    writeData(wb, sheet = "Introduction", intro_published_date, 
              startRow = 2, startCol = 1, colNames = FALSE)
    
    writeData(wb, sheet = "Introduction", guidance_hyperlink, 
              startRow = 5, startCol = 1, colNames = FALSE)
    
    writeData(wb, sheet = "Introduction", tools_hyperlink, 
              startRow = 10, startCol = 1, colNames = FALSE)
    
    writeData(wb, sheet = "Introduction", rights_hyperlink,
              startRow = 12, startCol = 1, colNames = FALSE)
    
    writeData(wb, sheet = "Data sheet", intro_title, 
              startRow = 1, startCol = 1, colNames = FALSE)
    
    writeData(wb, sheet = "Data sheet", guidance_content, 
              startRow = 3, startCol = 1, colNames = FALSE)
    
    writeData(wb, sheet = "Data sheet totals", datatotal_formulas,
              startRow = 1, startCol = 2, colNames = FALSE)
    
    # We can now delete all the labels (e.g. heading, text etc. )
    deleteData(wb, sheet = "Data sheet", 
               cols = 2, rows = str_which(guidance_content$rec_number, "Heading|Subheading|Subsubheading|Text")+2, 
               gridExpand = TRUE)
    
    # Add and style the update date as well if there is one
    if (length(guidance_info) == 3) {
        
        writeData(wb, sheet = "Introduction", intro_update_date, 
                  startRow = 3, startCol = 1, colNames = FALSE)
        
        addStyle(wb, sheet = "Introduction", intro_date_style, 
                 rows = 3, cols = 1, stack = FALSE)
    }
    
    
    ### Apply all of the appropriate styles ###
    
    addStyle(wb, sheet = "Introduction", intro_title_style, 
             rows = 1, cols = 1, stack = FALSE)
    
    addStyle(wb, sheet = "Introduction", intro_date_style, 
             rows = 2, cols = 1, stack = FALSE)
    
    addStyle(wb, sheet = "Introduction", hyperlink_style, 
             rows = 5, cols = 1, stack = FALSE)
    
    addStyle(wb, sheet = "Introduction", hyperlink_style, 
             rows = 10, cols = 1, stack = TRUE)
    
    addStyle(wb, sheet = "Introduction", hyperlink_style, 
             rows = 12, cols = 1, stack = FALSE)
    
    addStyle(wb, sheet = "Data sheet", datasheet_title_style, 
             rows = 1, cols = 1, stack = FALSE)
    
    addStyle(wb, sheet = "Data sheet", header_style, 
             rows = str_which(guidance_content$rec_number, "Heading|Subsubheading")+2, 
             cols = 1:13, stack = FALSE, gridExpand = TRUE)
    
    addStyle(wb, sheet = "Data sheet", subheader_style, 
             rows = str_which(guidance_content$rec_number, "Subheading")+2, 
             cols = 1:13, stack = FALSE, gridExpand = TRUE)
    
    addStyle(wb, sheet = "Data sheet", text_style, 
             rows = str_which(guidance_content$rec_number, "Heading|Subheading|Subsubheading|Text", negate = TRUE)+2, 
             cols = 1:13, stack = FALSE, gridExpand = TRUE)
    
    addStyle(wb, sheet = "Data sheet", highlight_style, 
             rows = which(guidance_content$rec_number == "Text")+2, 
             cols = 1:13, stack = FALSE, gridExpand = TRUE)
    
    addStyle(wb, sheet = "Data sheet", border_style, 
             rows = str_which(guidance_content$rec_number, "Heading|Subheading|Subsubheading", negate = TRUE)+2, 
             cols = 1:13, stack = TRUE, gridExpand = TRUE)
    
    
    ### Add in the drop downs and conditional formatting ###
    # the dataValidation arguments will show warnings as type list is used (known issue)
    
    # Yes/No/partial drop downs in column D and F
    dataValidation(wb, sheet = "Data sheet", 
                   rows = str_which(guidance_content$rec_number, "[:digit:]")+2, cols = c(4,6), 
                   type = "list", value = "'Dropdowns'!$A$1:$A$3")
    
    # Yes/No drop downs in column H
    dataValidation(wb, sheet = "Data sheet", 
                   rows = str_which(guidance_content$rec_number, "[:digit:]")+2, cols = 8, 
                   type = "list", value = "'Dropdowns'!$A$1:$A$2")
    
    #  Set up conditional formatting to grey out row when rec is not relevant (No selected in col D)
    conditionalFormatting(wb, sheet = "Data sheet", 
                          rows = str_which(guidance_content$rec_number, "[:digit:]")+2, cols = 5:12, 
                          rule = '$D5="No"', style = createStyle(bgFill = "#808080")) 
    
    ### Return the complete BAT ###
    
    return(wb)
}


