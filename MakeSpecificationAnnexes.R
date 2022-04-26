col_config <- 
  readr::read_csv(
    "~/Codes/Make-RGCS-Data-Audit-Guidance/inputs/2022-03-21 Column configuration OFFICIAL - Sheet 1 (4).csv"
  )

cols_included <- 
  col_config %>% 
  dplyr::filter(status == "Included")

sheet_title_style <- 
  openxlsx::createStyle(fontSize = 24, fontName = "Arial", textDecoration = "bold")

field_title_style <- 
  openxlsx::createStyle(fontName = "Arial", textDecoration = "bold")

body_style <- 
  openxlsx::createStyle(wrapText = T)

writeSheet <- function(col_id, status, field_number, col_long_name, definition, guidance_notes, purpose, in_gcs_strategy, sponsor, notes, completion_mode, form_regex, options, allowed_values, allowed_regex, numeric_range_min, numeric_range_max){
  sheet_name = paste0(field_number, ". ", col_id)
  
  openxlsx::addWorksheet(wb, sheet_name)
  
  sheet_text <- tibble::tibble(
    heading = c(paste0("Field ", field_number, " - ", col_long_name), "", "Field Number", "Name", "Abbreviation", "Definition", "Reason for Collection", "Notes"),
    content = c("", "", field_number, col_long_name, col_id, definition, purpose, guidance_notes)
  )
  
  sheet_text_length <- nrow(sheet_text)
  
  openxlsx::writeData(wb, sheet_name, sheet_text, colNames = F)
  
  openxlsx::addStyle(wb, sheet_name, sheet_title_style, rows = 1, cols = 1)
  
  openxlsx::setColWidths(wb, sheet_name, cols = c(1,2), widths = c(24, 130))
  
  openxlsx::mergeCells(wb, sheet_name, cols = 1:2, rows = 1)
  
  openxlsx::addStyle(wb, sheet_name, field_title_style, cols = 1, rows = c(3, 4, 5, 6, 7, 8))
  
  openxlsx::addStyle(wb, sheet_name, body_style, cols = 2, rows = c(3, 4, 5, 6, 7, 8))
  
  if(completion_mode == "discrete_choices"){
    
    item_labels = unlist(stringr::str_split(string = options, pattern = ";"))
    
    extra_list_items = c("Other", "Unknown")
    
    `Item Label` <- c(item_labels, extra_list_items)
    
    ## Generate item codes to match the labels
    item_codes <- 10:(10+length(item_labels)-1)
    
    extra_item_codes <- c(99, "**")
    
    `Item Code` <- c(item_codes, extra_item_codes)
    
    item_table <- tibble::tibble(
      `Item Code`,
      `Item Label`
    )
    
    item_table_start_row = sheet_text_length + 2
      
    openxlsx::writeData(wb, sheet_name, item_table, startRow = item_table_start_row)
    
  } else if (completion_mode %in% c("free_text", "restricted_text", "discrete_choices_large", "continuous_number")){
    
    item_labels = unlist(stringr::str_split(string = options, pattern = ";"))
    
    extra_list_items = c("Unknown")
    
    `Description` <- c(item_labels, extra_list_items)
    
    item_codes <- unlist(stringr::str_split(string = allowed_values, pattern = ";"))
    
    extra_item_codes <- c("**")
    
    `Contents` <- c(item_codes, extra_item_codes)
    
    item_table <- tibble::tibble(
      `Contents`,
      `Description`
    )
    
    item_table_start_row = sheet_text_length + 2
    
    openxlsx::writeData(wb, sheet_name, item_table, startRow = item_table_start_row)
    
  }
  
  openxlsx::addStyle(wb, sheet_name, field_title_style, cols = c(1, 2), rows = item_table_start_row)
  
}

wb <- openxlsx::createWorkbook()

cols_included %>%
  purrr::pmap(
    .f = writeSheet
  )

openxlsx::saveWorkbook(wb, "~/Codes/Make-RGCS-Data-Audit-Guidance/outputs/2022-04-24 GCS Data Audit 2022 - Specification Annexes OFFICIAL.xlsx", overwrite = T)
