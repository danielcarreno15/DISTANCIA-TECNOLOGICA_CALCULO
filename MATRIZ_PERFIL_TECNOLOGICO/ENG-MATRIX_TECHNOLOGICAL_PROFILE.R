# Load necessary packages
install.packages("readxl")
install.packages("writexl")
install.packages("tidyverse")  # Includes tidyr and dplyr
install.packages("stringr")
install.packages("openxlsx")

library(readxl)
library(writexl)
library(tidyr)
library(dplyr)
library(stringr)  
library(openxlsx)

# Specify the file path
file.choose()

# Specify the path of the file
excel_path <- "C:\\Users\\danie\\Downloads\\resultList (1).xlsx"

# Read the Excel file
data <- read_excel(excel_path)

# Verify that data is a data frame
if (!is.data.frame(data)) {
  stop("The object data is not a data frame")
}

# Display the first few records to check the structure
print(head(data))

# Verify column names
print(colnames(data))

# Ensure that the "Application Date" column exists
if (!"Application Date" %in% colnames(data)) {
  stop("The 'Application Date' column does not exist in the data frame")
}

# Convert Application Date to date format
data <- data %>%
  mutate(Application_Date = as.Date(`Application Date`, format = "%d.%m.%Y"))

# Function to filter by year range and count patents
filter_and_count <- function(data, start_year, end_year) {
  filtered_data <- data %>%
    filter(format(Application_Date, "%Y") >= start_year & format(Application_Date, "%Y") <= end_year)
  
  count_per_year <- filtered_data %>%
    group_by(year = format(Application_Date, "%Y")) %>%
    summarize(count = n())
  
  total_patents <- nrow(filtered_data)
  range <- paste(start_year, end_year, sep = "-")
  
  list(data = filtered_data, count_per_year = count_per_year, total_patents = total_patents, range = range)
}

# Filter by specified ranges
range_1 <- filter_and_count(data, 2012, 2016)
range_2 <- filter_and_count(data, 2013, 2017)
range_3 <- filter_and_count(data, 2014, 2018)

# HOW TO FIND TECHNOLOGICAL DISTANCE

# Create a new data frame with the summarized results
df_range_1 <- range_1$data
df_range_2 <- range_2$data
df_range_3 <- range_3$data

# Check if the "Application Id" column exists
if(!"Application Id" %in% names(df_range_1)) {
  stop("The 'Application Id' column does not exist in df_range_1")
}
if(!"Application Id" %in% names(df_range_2)) {
  stop("The 'Application Id' column does not exist in df_range_2")
}
if(!"Application Id" %in% names(df_range_3)) {
  stop("The 'Application Id' column does not exist in df_range_3")
}

# Function to separate IPC codes into columns and keep the other columns
separate_codes <- function(df) {
  # Separate IPC codes into columns
  df_codes <- df %>%
    separate_rows(`I P C`, sep = ";") %>%
    mutate(`I P C` = str_trim(`I P C`))
  
  # Extract the first three digits of each IPC code
  df_codes <- df_codes %>%
    mutate(Code = substr(`I P C`, 1, 3))
  
  # Combine separated IPC codes into a single column for each "Application Id"
  df_codes_combined <- df_codes %>%
    group_by(`Application Id`) %>%
    summarise(Unique_Codes = paste(unique(Code), collapse = ";")) %>%
    ungroup()
  
  # Combine the original DataFrame with the separated codes
  df_final <- df %>%
    select(-`I P C`) %>%
    left_join(df_codes_combined, by = "Application Id")
  
  return(df_final)
}

# Apply the function to the DataFrames
df_range_1_levels <- separate_codes(df_range_1)
df_range_2_levels <- separate_codes(df_range_2)
df_range_3_levels <- separate_codes(df_range_3)

# Function to create a comparison matrix with a summation column
create_comparison_matrix <- function(df_levels, target_codes) {
  # Get the names of the columns containing the IPC codes
  code_columns <- grep("Unique_Codes", names(df_levels), value = TRUE)
  
  # Get the Application Ids
  application_ids <- df_levels$`Application Id`
  
  # Create an empty matrix with the appropriate dimensions
  matrix <- matrix(0, nrow = length(target_codes), ncol = nrow(df_levels))
  
  # Fill the matrix with 1 if there is a match for IPC codes
  for (i in 1:nrow(df_levels)) {
    present_codes <- str_split(df_levels[i, code_columns], ";") %>% unlist() %>% na.omit()
    for (j in 1:length(target_codes)) {
      if (target_codes[j] %in% present_codes) {
        matrix[j, i] <- 1
      }
    }
  }
  
  # Convert the matrix to a data frame
  df_matrix <- as.data.frame(matrix)
  
  # Add column names (Application Ids)
  colnames(df_matrix) <- application_ids
  
  # Add the codes as a column to the data frame
  df_matrix <- cbind(Code = target_codes, df_matrix)
  
  # Check if df_matrix has more than one column after excluding the "Code" column
  if (ncol(df_matrix) > 2) {  # If it has more than two columns (including "Code" and at least one data column)
    # Add a column with the summation per row
    df_matrix$Total <- rowSums(df_matrix[, -1])  # Exclude the "Code" column in the sum calculation
  } else {
    # If there is only one data column, the sum for that row is simply the value of that column
    df_matrix$Total <- df_matrix[, 2]
  }
  
  return(df_matrix)
}

# Define the target codes
target_codes <- c(
  "A01", "A21", "A22", "A23", "A24", "A41", "A42", "A43", "A44",  
  "A45", "A46", "A47", "A61", "A62", "A63", "A99", "B01", "B02", "B03", "B04",  
  "B05", "B06", "B07", "B08", "B09", "B21", "B22", "B23", "B24", "B25", "B26",  
  "B27", "B28", "B29", "B30", "B31", "B32", "B33", "B41", "B42", "B43",  
  "B44", "B60", "B61", "B62", "B63", "B64", "B65", "B66", "B67", "B68", "B81",  
  "B82", "B99", "C01", "C02", "C03", "C04", "C05", "C06", "C07", "C08", "C09",  
  "C10", "C11", "C12", "C13", "C14", "C21", "C22", "C23", "C25", "C30", "C40",  
  "C99", "D01", "D02", "D03", "D04", "D05", "D06", "D07", "D21", "D99", "E01",  
  "E02", "E03", "E04", "E05", "E06", "E21", "E99", "F01", "F02", "F03", "F04",  
  "F15", "F16", "F17", "F21", "F22", "F23", "F24", "F25", "F26", "F27", "F28",  
  "F41", "F42", "F99", "G01", "G02", "G03", "G04", "G05", "G06", "G07", "G08",  
  "G09", "G10", "G11", "G12", "G16", "G21", "G99", "H01", "H02", "H03", "H04",  
  "H05", "H10", "H99") 

# Create comparison matrices
matrix_range_1 <- create_comparison_matrix(df_range_1_levels, target_codes)
matrix_range_2 <- create_comparison_matrix(df_range_2_levels, target_codes)
matrix_range_3 <- create_comparison_matrix(df_range_3_levels, target_codes)

# Create a new Excel file and add sheets with the results
wb <- createWorkbook()

add_sheet <- function(wb, range, sheet_name) {
  addWorksheet(wb, sheet_name)
  if (!is.null(range$data) && nrow(range$data) > 0) {
    writeData(wb, sheet_name, range$data, startCol = 1, startRow = 1)
    writeData(wb, sheet_name, paste("Total patents in the range", range$range, ":", range$total_patents), startCol = 1, startRow = nrow(range$data) + 2)
    if (!is.null(range$count_by_year) && nrow(range$count_by_year) > 0) {
      writeData(wb, sheet_name, range$count_by_year, startCol = 1, startRow = nrow(range$data) + 4)
    }
  } else {
    writeData(wb, sheet_name, "No data available", startCol = 1, startRow = 1)
  }
}

# Add sheets with the filtered data
add_sheet(wb, range_1, "2012-2016")
add_sheet(wb, range_2, "2013-2017")
add_sheet(wb, range_3, "2014-2018")

# Add each matrix as a separate sheet
addWorksheet(wb, "Matrix_Range_1")
writeData(wb, "Matrix_Range_1", matrix_range_1)

addWorksheet(wb, "Matrix_Range_2")
writeData(wb, "Matrix_Range_2", matrix_range_2)

addWorksheet(wb, "Matrix_Range_3")
writeData(wb, "Matrix_Range_3", matrix_range_3)

# Save the Excel file
saveWorkbook(wb, "C:\\Users\\danie\\Documents\\EMPRESAS DT\\EMPRESA_EJEMPLO\\EMPRESA 4.xlsx", overwrite = TRUE)

# Function to process each Excel file for storage
process_and_store_excel <- function(excel_path, storage_path) {
  # Read the generated Excel file
  data_matrix_1 <- read_excel(excel_path, sheet = "Matrix_Range_1")
  data_matrix_2 <- read_excel(excel_path, sheet = "Matrix_Range_2")
  data_matrix_3 <- read_excel(excel_path, sheet = "Matrix_Range_3")
  
  # Check that data_matrix are data frames
  if (!is.data.frame(data_matrix_1) | !is.data.frame(data_matrix_2) | !is.data.frame(data_matrix_3)) {
    stop("One of the data objects is not a data frame")
  }
  
  # Extract IPC codes from the matrix sheets
  ipc_codes_1 <- data_matrix_1$Code
  ipc_codes_2 <- data_matrix_2$Code
  ipc_codes_3 <- data_matrix_3$Code
  
  # Extract the "Total" column from the matrix sheets
  total_1 <- data_matrix_1$Total
  total_2 <- data_matrix_2$Total
  total_3 <- data_matrix_3$Total
  
  # Create a new data frame with IPC codes as columns
  df_result_1 <- tibble(
    Company = tools::file_path_sans_ext(basename(excel_path)),
    !!!setNames(as.list(total_1), ipc_codes_1)
  )
  
  df_result_2 <- tibble(
    Company = tools::file_path_sans_ext(basename(excel_path)),
    !!!setNames(as.list(total_2), ipc_codes_2)
  )
  
  df_result_3 <- tibble(
    Company = tools::file_path_sans_ext(basename(excel_path)),
    !!!setNames(as.list(total_3), ipc_codes_3)
  )
  
  # Read the current content of the storage file
  if (file.exists(storage_path)) {
    wb_storage <- loadWorkbook(storage_path)
  } else {
    # If the file does not exist, create it and add the necessary sheets
    wb_storage <- createWorkbook()
    addWorksheet(wb_storage, "Storage_Range_1")
    addWorksheet(wb_storage, "Storage_Range_2")
    addWorksheet(wb_storage, "Storage_Range_3")
  }
  
  # Write the data to each sheet (checking if they already exist)
  if ("Storage_Range_1" %in% names(wb_storage)) {
    current_data_1 <- read.xlsx(wb_storage, sheet = "Storage_Range_1")
    writeData(wb_storage, "Storage_Range_1", 
              rbind(current_data_1, df_result_1), colNames = TRUE)
  }
  
  if ("Storage_Range_2" %in% names(wb_storage)) {
    current_data_2 <- read.xlsx(wb_storage, sheet = "Storage_Range_2")
    writeData(wb_storage, "Storage_Range_2", 
              rbind(current_data_2, df_result_2), colNames = TRUE)
  }
  
  if ("Storage_Range_3" %in% names(wb_storage)) {
    current_data_3 <- read.xlsx(wb_storage, sheet = "Storage_Range_3")
    writeData(wb_storage, "Storage_Range_3", 
              rbind(current_data_3, df_result_3), colNames = TRUE)
  }
  
  # Save the storage Excel file
  saveWorkbook(wb_storage, storage_path, overwrite = TRUE)
}
# Path of the storage Excel file
storage_path <- "C:\\Users\\danie\\Documents\\EMPRESAS DT\\EMPRESA_EJEMPLO\\STORAGE_DT.xlsx"

# Process and store the information from the generated file
process_and_store_excel("C:\\Users\\danie\\Documents\\EMPRESAS DT\\EMPRESA_EJEMPLO\\EMPRESA 4.xlsx", storage_path)

# Get the current working directory
print(getwd())
