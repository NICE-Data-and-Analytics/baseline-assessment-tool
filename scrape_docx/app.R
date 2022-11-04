library(tidyverse)
library(shiny)
library(openxlsx)
library(officer)
library(shinyjs)

# Load functions, defined in separate script
source("./scrape_docx_fn.R")
source("create_BAT.R")

# Define UI 
ui <- fluidPage(
    
    # Shinyjs is needed to disable the download button on app launch, so users know to upload a file first
    useShinyjs(),
    
    # Change app theme
    theme = bslib::bs_theme(bootswatch = "darkly"),
    
    # Application title
    titlePanel("Generating BATs"),
    
    tabsetPanel(
        tabPanel("From Word doc",
                 # Only used to name the output file
                 textInput("guideline_number",
                           "1. What is the guideline number?",
                           placeholder = "e.g. NG205 (To name output file.)"),
                 fileInput("file", 
                           "2. Upload DOCX file with guideline", 
                           accept = "docx"),
                 p("3. Download Excel spreadsheet with extracted recommendations"),
                 disabled(
                     downloadButton("download",
                                "Download file")
                 ),
                 br(),
                 br(),
                 p("4. Copy and paste from downloaded spreadsheet into BAT template."),
                 tags$ul(
                     tags$li("Check for errors in the extracted text"),
                     tags$li("Check hyphenated words, e.g. 'omega-3' - the hyphen is often lost when pulled"),
                     tags$li(HTML("Add <b>bold formatting</b> to the text as needed - formatting is not preserved in the extraction")),
                     tags$li("Re-insert paragraph spacing (new line) for recommendations with multiple paragraphs (not bullet points) - currently merged into one big paragraph when pulled")
                 )
        ),
        tabPanel("From website",
                 HTML("<p>Download the full recommendation set from <a href='https://norma.nice.org.uk'>NORMA (NICE-ONS Recommendation Matching Algorithm)</a>.</p>"),
                 p("Note, NORMA:"), 
                 tags$ul(
                    tags$li("is updated once a day"),
                    tags$li("can only be accessed from the office or using VMware"),
                    tags$li("merges bulletpoints into a big chunk of text")
                    ),
                 HTML("<p>See <a href='https://space.nice.org.uk/sorce/beacon/singlepageview.aspx?pii=1895&row=26761'>this blog</a> for more info on NORMA.</p>")
        )
    )
)

# Define server function
server <- function(input, output, session) {
    
    # Read the uploaded DOCX file
    docx <- reactive({
        req(input$file)
        
        read_docx(input$file$datapath)
        
    })
    
    # Run the custom function scrape_docx to extract recommendatioms
    # Returns a list with the extracted table (recs()$table) and 
    # an Excel spreadsheet with the table (docs()$wb) - this is a Workbook object, part of the openxlsx package
    recs <- reactive(scrape_docx(docx()))
    
    
    # Run the custom function scrape_title_and_dates to get the guidance title,
    # publication date, and update date
    title_dates <- reactive(scrape_title_and_dates(docx()))
    
    # Run the custom function create_BAT to create the BAT
    completed_BAT <- reactive(create_BAT(guidance_number = input$guideline_number,
                                         guidance_info = title_dates(),
                                         guidance_content = recs()))
    
    # Enable download button once a file has been uploaded
    observeEvent(input$file, {
        enable("download")
    })
    
    # Download button
    output$download <- downloadHandler(
        filename = function() {
            # Generate file name using user-inputted guideline number
            paste0(input$guideline_number, "_recs", ".xlsx")
        },
        content = function(file) {
            # Save the Workbook object to a file
            saveWorkbook(completed_BAT(), file)
        }
    )
    
}

# Run the application
shinyApp(ui = ui, server = server)

