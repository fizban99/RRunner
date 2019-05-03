# Helper functions to read the csv files generated from Excel
library(readxl)
library(writexl)

getTable <- function(tableName) {
  read_excel("../tmp/_RInput_.xlsx", sheet = tableName)
}


writeResult <- function(tablenames) {
  write_xlsx(tablenames, path = "../tmp/_ROutput_.xlsx", col_names = TRUE, format_headers = FALSE)
}


done <- function() {
  file.create("../tmp/done")
  closeAllConnections()
}