# Shiny dynamic module excel example

This repository contains an shiny app that shows how to call shiny modules  depending on external conditions like the number of excel file the user opens or the number of sheets inside those excel files.
Shiny's new moduleServer is used instead of callModule. 

## Applications

### `Shiny dynamic module Excel example` 
App to explore Excel files and their sheets by uploading them and displaying the data sheets in data tables tab by tab. 
The app creates two example Excel files in the current  working directory ExcelFile1.xlsx and ExcelFile2.xlsx with multiple sheets containing datasets.
You can open this files or your own once for testing.

__Highlights:__

- `shiny::{moduleServer()}`
- `purrr::{pmap(), pmap_chr()}`
- `openxlsx::{read.xlsx(), loadWorkbook()}`
  
