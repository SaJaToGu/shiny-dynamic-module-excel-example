#
# This is a Shiny web application. You can run the application by clicking
# the 'Run App' button above.
#
# Find out more about building applications with Shiny here:
#
#    http://shiny.rstudio.com/
#

library(shiny)
library(tidyverse)
library(openxlsx)

# Create two excel files with data sets as examples 
wb <- openxlsx::createWorkbook()
openxlsx::addWorksheet(wb, "Cars")
openxlsx::writeDataTable(wb, "Cars", datasets::mtcars, tableName = "cars")
openxlsx::addWorksheet(wb, "Iris")
openxlsx::writeDataTable(wb, "Iris", datasets::iris, tableName = "iris")
openxlsx::addWorksheet(wb, "Faithful")
openxlsx::writeDataTable(wb, "Faithful", datasets::faithful, tableName = "faithful")
xlsxFile <- "./ExcelFile1.xlsx"
openxlsx::saveWorkbook(wb, file = xlsxFile, overwrite = TRUE)
wb <- openxlsx::createWorkbook()
openxlsx::addWorksheet(wb, "pressure")
openxlsx::writeDataTable(wb, "pressure", datasets::pressure, tableName = "pressure")
openxlsx::addWorksheet(wb, "quakes")
openxlsx::writeDataTable(wb, "quakes", datasets::quakes, tableName = "quakes")
openxlsx::addWorksheet(wb, "rock")
openxlsx::writeDataTable(wb, "rock", datasets::rock, tableName = "rock")
openxlsx::addWorksheet(wb, "warpbreaks")
openxlsx::writeDataTable(wb, "warpbreaks", datasets::warpbreaks, tableName = "warpbreaks")
xlsxFile <- "./ExcelFile2.xlsx"
openxlsx::saveWorkbook(wb, file = xlsxFile, overwrite = TRUE)


# Define ExcelSheetUi module for application that shows one excel sheet
ExcelSheetUi <-
    function(id,
             title = id,
             ...,
             value = title,
             icon = NULL) {
        ns <- shiny::NS(id)

        shiny::tabPanel(
            title,
            DT::dataTableOutput(outputId = ns("sheet_data"))
        )
    }

# Define ExcelSheetSrv module for application that shows one excel sheet
ExcelSheetSrv <-
    function(id,
             workbook,
             sheet) {
        moduleServer(
            id,
            function(input, output, session) {
                ns <- session$ns
                workbook_re <- shiny::reactive({
                    if (shiny::is.reactive(workbook)) {
                        result <- workbook()
                    } else {
                        result <- workbook
                    }
                    return(result)
                })
                
                rawData_re <- shiny::reactive({
                    shiny::req(workbook_re())
                    openxlsx::read.xlsx(workbook_re(),
                                        sheet = sheet)
                })
                
                output$sheet_data <- DT::renderDataTable({
                    shiny::req(rawData_re())
                    DT::datatable(
                        rawData_re(),
                        selection = 'none',
                        filter = 'top',
                        extensions = c("Buttons"),
                        options = list(
                            pageLength = 100,
                            lengthMenu = list(c(5, 10, 50, 100,-1),
                                              c('5', '10', '50', '100', 'All')),
                            dom = 'lBfrtip',
                            buttons = c('excel', 'copy', 'csv', 'pdf', 'print', 'colvis')
                        )
                    )
                })
            }
        )
    }

# Define ExcelUI module for application that shows an excel file with all it's sheets
ExcelUi <-
    function(id,
             title = id,
             ...,
             value = title,
             icon = NULL) {
        ns <- shiny::NS(id)
        shiny::tabPanel(title = title,
                        shiny::tabsetPanel(id = ns("Sheets")
                 )
        )
    }

# Define ExcelSrv module for application that shows an excel file with all it's sheets
ExcelSrv <-
    function(id, ExcelFile = NULL) {
        moduleServer(
            id,
            function(input, output, session) {
                
                ns <- session$ns
                
                ExcelFile_re <- shiny::reactive({
                    if (shiny::is.reactive(ExcelFile)) {
                        result <- ExcelFile()
                    } else {
                        result <- ExcelFile
                    }
                    return(result)
                })
                
                workbook_re <- reactive({
                    shiny::req(ExcelFile_re())
                    openxlsx::loadWorkbook(ExcelFile_re()$datapath)
                })
                
                
                # Upload ---------------------------------------------------------------
                shiny::observeEvent(ExcelFile_re,{
                    # uncomment following line for debugging
                    # browser()
                    # prepare excel sheet UIs
                    sheetlist <- openxlsx::getSheetNames(ExcelFile_re()$datapath) %>%
                        tidyr::as_tibble() %>%
                        dplyr::rename(sheetname = value) %>%
                        dplyr::mutate(sheetid = purrr::pmap_chr(., function(sheetname, ...) make.names(sheetname))) %>%
                        dplyr::mutate(nssheetid = purrr::pmap_chr(., function(sheetid, ...) ns(sheetid))) %>%
                        dplyr::mutate(taglist = purrr::pmap(., function(nssheetid, sheetname, ...) ExcelSheetUi(id = nssheetid, title = sheetname)))
                    
                    # UI create a new tab for each excel sheet (dynamic)
                    sheetlist %>%
                        mutate(is_first = (sheetname == first(sheetname))) %>% 
                        purrr::pmap(., function(taglist, is_first, ...) shiny::appendTab(inputId = "Sheets",
                                                                               taglist,
                                                                               select = is_first))
                    # Server call module for each excel sheet (dynamic)
                    sheetlist %>%
                        purrr::pmap(., function(sheetid, sheetname, ...) ExcelSheetSrv(
                            id = sheetid,
                            workbook = workbook_re,
                            sheet = sheetname))
                })
            })
    }

# Define ShinyDynModExcelUi for module that uploads excel files and shows their sheets in tabs
ShinyDynModExcelUi <-
    function(id,
             title = id,
             ...,
             value = title,
             icon = NULL) {
        ns <- shiny::NS(id)
        shiny::tabPanel(title = title,
                        shiny::fileInput(ns("file"), "Upload Excel files", buttonLabel = "Upload..."),
                        shiny::actionButton(inputId = ns("Close"), label = "Close Excel"),
                        shiny::tabsetPanel(id = ns("excels"))
        )
    }

# Define ShinyDynModExcelSrv for module for that shows excel files with all it's sheets
ShinyDynModExcelSrv <-
    function(id) {
        moduleServer(
            id,
            function(input, output, session) {
                ns <- session$ns
                # Close ---------------------------------------------------------------
                observeEvent(input$Close,{
                    shiny::removeTab(inputId = "excels", target = input$excels)
                })
                # Upload Excel files ---------------------------------------------------------------
                shiny::observeEvent(input$file,{
                    shiny::req(input$file$name)
                    ExcelTabId <- make.names(input$file$name)
                    ExcelTabNsId <- ns(make.names(input$file$name))
                    ExcelTabTitle <- input$file$name
                    shiny::removeTab(inputId = "excels", ExcelTabId)
                    shiny::appendTab(inputId = "excels",
                                     ExcelUi(id = ExcelTabNsId, title = ExcelTabTitle),
                                     select = TRUE)
                    ExcelSrv(id = ExcelTabId, ExcelFile = input$file)
                })
            })
    }

ui <- shinyUI(
    navbarPage(
        "Shiny dynamic module Excel example",
        id = "ShinyDynModExcelEample",
        tabPanel("Excel Files", ShinyDynModExcelUi("ShinyDynModExcel")),
        tabPanel(
            "Any questions or comments can be sent to",
            br(),
            "Guido Berning: " ,
            a("guido.berning@covestro.com", href = "mailto:guido.berning@covestro.com")
        )
    )
)

server <- shinyServer(function(input, output, session) {
    ShinyDynModExcelSrv(id = "ShinyDynModExcel")
})

# Run the application 
shinyApp(ui = ui, server = server)
