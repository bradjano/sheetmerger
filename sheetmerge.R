library("tidyverse")
library("readxlsx")
library("openxlsx")

# Loads the XLSX as a workbook object
wb <- loadWorkbook("Paradata.xlsx")

# Create a list of sheet names for the workbook object
mysheets <- sheets(wb)

# Function to import, mutate, and save workbook sheetss
intake <- function(sheetna){
  csvna <- str_c(sheetna,".csv")
  path <- str_c("CSVs/",csvna)
  d <- read_xlsx("Paradata.xlsx", sheet = sheetna)
  d <- d %>% 
    mutate(cluster = as.numeric(sheetna)) %>% 
    write_csv(path)
}

# Loop through sheets and apply intake function
for (i in mysheets) {
  intake(i)
}

# List, import, and combine all CSVs in the folder
clusters <- list.files("CSVs", pattern = "*.csv", full.names = TRUE) %>% 
  lapply(read_csv) %>% 
  bind_rows() %>% 
  write_csv("Paradata_clusters.csv")
