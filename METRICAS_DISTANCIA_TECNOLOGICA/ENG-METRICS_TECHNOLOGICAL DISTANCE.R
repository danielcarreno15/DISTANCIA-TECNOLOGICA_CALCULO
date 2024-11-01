# Load necessary packages
install.packages("readxl")
install.packages("writexl")
install.packages("tidyverse")  # This includes tidyr and dplyr
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

# Specify the file path
excel_path <- "C:\\Users\\danie\\Documents\\EMPRESAS DT\\EMPRESA_EJEMPLO\\STORAGE_DT.xlsx"

# Read data from each sheet
profiles1 <- read_excel(excel_path, sheet = "Storage_Range_1")
profiles2 <- read_excel(excel_path, sheet = "Storage_Range_2")
profiles3 <- read_excel(excel_path, sheet = "Storage_Range_3")

# Check if there are NA or Inf values in the data
check_valid_data <- function(df) {
  return(!any(is.na(df)) && all(is.finite(df)))
}
# Normalize data by row (proportions)
normalize_data <- function(df) {
  return(as.data.frame(t(apply(df[, -1], 1, function(x) x / sum(x, na.rm = TRUE)))))
}

profiles1_normalized <- cbind(profiles1[1], normalize_data(profiles1))
profiles2_normalized <- cbind(profiles2[1], normalize_data(profiles2))
profiles3_normalized <- cbind(profiles3[1], normalize_data(profiles3))

# Function to calculate Jaffe distance (adjusted)
jaffe_distance <- function(p1, p2) {
  return(1 - (sum(p1 * p2) / sqrt(sum(p1^2) * sum(p2^2))))
}

# Function to calculate Euclidean distance
euclidean_distance <- function(p1, p2) {
  return(sqrt(sum((p1 - p2)^2)))
}

# Function to calculate Minimum Complement distance
min_complement_distance <- function(p1, p2) {
  sum_pmin <- sum(pmin(p1, p2))
  return(1 - sum_pmin)
}

# Function to calculate distances between two specific companies in a sheet
calculate_specific_distances <- function(profiles, company1, company2) {
  company_names <- profiles$Company
  row1 <- which(company_names == company1)
  row2 <- which(company_names == company2)
  
  if(length(row1) == 0 || length(row2) == 0) {
    stop("One or both companies are not found in the data.")
  }
  p1 <- as.numeric(profiles[row1, -1])
  p2 <- as.numeric(profiles[row2, -1])
  
  dist_jaffe <- jaffe_distance(p1, p2)
  dist_euclidean <- euclidean_distance(p1, p2)
  dist_min_complement <- min_complement_distance(p1, p2)
  
  return(list(jaffe = dist_jaffe, euclidean = dist_euclidean, min_complement = dist_min_complement))
}

# Calculate the distances between the two specific companies for each profile
company1 <- "EMPRESA 3"
company2 <- "EMPRESA 4"

distances1 <- calculate_specific_distances(profiles1_normalized, company1, company2)
distances2 <- calculate_specific_distances(profiles2_normalized, company1, company2)
distances3 <- calculate_specific_distances(profiles3_normalized, company1, company2)

# Display the results
print(paste("Distances in storage_range_1:"))
print(distances1)
print(paste("Distances in storage_range_2:"))
print(distances2)
print(paste("Distances in storage_range_3:"))
print(distances3)

# Save the results to Excel files
results <- data.frame(
  Profile = c("storage_range_1", "storage_range_2", "storage_range_3"),
  Jaffe = c(distances1$jaffe, distances2$jaffe, distances3$jaffe),
  Euclidean = c(distances1$euclidean, distances2$euclidean, distances3$euclidean),
  Min_Complement = c(distances1$min_complement, distances2$min_complement, distances3$min_complement)
)

write_xlsx(list("Results" = results), "DT_EMPRESA.xlsx")
