# Helper functions to read the csv files generated from Excel
library(readxl)
library(writexl)

getTable <- function(tableName) {
  read_excel("_Input_.xlsx", sheet = tableName)
}


writeResult <- function(tablenames) {
  write_xlsx(tablenames, path = "_Output_.xlsx", col_names = TRUE, format_headers = FALSE)
}


done <- function() {
  file.create("done")
  closeAllConnections()
}