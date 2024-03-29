scrape_docx <- function(doc) {
    # Reads the uploaded Word doc and returns a data.frame splitting out all the elements of the document
    content <- docx_summary(doc)
    
    # Identify row number where the recommendations start
    # Specifically, this picks up the first section heading, e.g. "1.1 Safety"
    rec_start <-  which(str_detect(content$style_name, "Numbered heading 2"))[[1]]
    
    # Where recommendations end
    rec_end <- which(str_detect(content$text, "Terms used in this guideline") & str_detect(content$style_name, "heading [12]")) - 1    
    
    if (is_empty(rec_end)) {
        rec_end <- which(str_detect(content$text, "Recommendations for research") & content$style_name == "heading 1") - 1
        if(is_empty(rec_end)) {
            rec_end <- which(str_detect(content$text, "Update information") & content$style_name == "heading 1") - 1
        }
    }
    
    # Manipulate text
    recommendations <- content %>% 
        # Only keep rows with recommendations
        slice(rec_start:rec_end) %>% 
        # Replace empty strings with NAs
        mutate(text = na_if(text, "")) %>% 
        # Remove panels, which contain text on the evidence and rationale behind recs
        filter(!(style_name %in% c("Panel (Default)", "NICE normal single spacing")),
               # Remove rows with no text
               !is.na(text),
               # Remove table cells
               content_type != "table cell") %>% 
        # Drop unnecessary columns
        select(!level:last_col()) %>%
        # docx_summary() function splits out the last bullet point as a different style, for some reason, e.g. "Bullet indent 1 last"
        # Some bullet indent styles also named differently, e.g. "Bullet indent 1 shaded"
        # Unify styles
        mutate(style_name = case_when(str_starts(style_name, "Bullet indent 1") | str_starts(style_name, "Bullet left") ~ "Bullet indent 1",
                                      str_starts(style_name, "Bullet indent 2") ~ "Bullet indent 2",
                                      # For updates, some recs given different style
                                      style_name == "Recommendation not updated" ~ "Numbered level 3 text",
                                      TRUE ~ style_name)) %>%
        group_by(style_name) %>% 
        # Number the sections
        # For rows which aren't section headers, give it an NA instead
        mutate(section = if_else(style_name == "Numbered heading 2", 
                                 1:n(), 
                                 NA_integer_)) %>% 
        ungroup() %>% 
        # Fill down so all the rows in a section have the right section number
        # Like dragging the bottom right corner down in Excel
        fill(section) %>% 
        # Format the section number as 1.1, 1.2 etc
        mutate(section = str_c("1", section, sep = ".")) %>% 
        group_by(section, style_name) %>% 
        # Number the recommendations, e.g. 1.1.5
        # All the recommendation paragraphs are in the 'Numbered level 3 text' style
        # Bulletpoints are in separate rows for now under a different style name
        mutate(rec_number = if_else(str_detect(style_name, "Numbered level 3|4 text"), 
                                    str_c(section, 1:n(), sep = "."),
                                    NA_character_)) %>% 
        ungroup() %>%
        # Fill down the rec number
        fill(rec_number) %>%
        # For non-recommendation text, based on the text style, specify heading, subheading etc in the rec number
        # For easy identification of non-rec text when copy and pasting
        mutate(rec_number = case_when(style_name == 'Numbered heading 2' ~ "Heading",
                                      str_detect(style_name, 'heading 3') ~ "Subheading",
                                      style_name == 'heading 4' ~ "Subsubheading",
                                      style_name == 'NICE normal' | str_starts(style_name, "Bullet left") ~ "Text",
                                      style_name == "Panel (Primary)" ~ "Panel",
                                      style_name == "caption" ~ "Caption",
                                      TRUE ~ rec_number)) %>% 
        # Add a bullet point symbol to the beginning of the string for bullet point text
        mutate(text = case_when(style_name == "Bullet indent 1" ~ paste0("\u2022 ", text),
                                style_name == "Bullet indent 2" ~ paste0("    - ", text),
                                TRUE ~ text)) %>% 
        group_by(rec_number) %>% 
        # Merge the bullet points under each recommendation (which are currently in separate rows) with the recommendation 
        mutate(text = if_else(style_name %in% c("Bullet indent 1", 
                                                "Numbered level 3 text", 
                                                "Numbered level 4 text",
                                                "Bullet indent 2",
                                                # In some docs, this is the style for additional paragraphs of the rec below bullet points
                                                "NICE normal indented"),
                              paste(text, collapse = "\n"),
                              text) %>% str_trim()) %>% 
        # Drop the bullet point and additional paragraph rows (which have still remained as separate rows) from the merge above
        filter(!str_starts(style_name, "Bullet indent"),
               style_name != "NICE normal indented") %>% 
        # Extract the year of publishing/updating into a separate column
        mutate(rec_year = str_extract(text, "(?<=\\[)\\d{4}.*(?=\\])") %>% 
                   str_remove("\\[|\\]"), 
               text = str_remove(text, "\\[\\d{4}.*\\]$") %>% 
                   str_trim()) %>% 
        # Remove hyperlinks
        mutate(text = str_remove_all(text, ' HYPERLINK(  \\\\l)? "[:graph:]+" ')) %>%
        # Remove REF thing
        mutate(text = str_remove_all(text, ' REF.+MERGEFORMAT ') %>% str_trim()) %>% 
        # Add heading number
        mutate(text = if_else(rec_number == "Heading", 
                              paste(section, text, sep = " "), 
                              text))

  # Create a Workbook object
  wb <- createWorkbook()

  # Add a worksheet
  addWorksheet(wb, "main")

  # Create a ready for export table
  recs_export <- recommendations %>%
    select(text, rec_number, rec_year)

  # Write the table into the worksheet
  writeDataTable(wb, "main", recs_export)

  # Make the cells have text wrapping
  addStyle(wb, "main",
    createStyle(wrapText = TRUE),
    rows = 1:nrow(recs_export),
    cols = 1:ncol(recs_export),
    gridExpand = TRUE
  )

  # Make the first column much wider
  setColWidths(wb, "main", cols = 1, widths = 60)

  # Return the recs table and Workbook object
  return(list(
    table = recommendations,
    wb = wb
  ))
}
