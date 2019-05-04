# Helper functions to read the csv files generated from Excel
library(readxl)
library(writexl)

getTable <- function(tableName) {
  read_excel("../tmp/_RInput_.xlsx", sheet = tableName)
}


writeResult <- function(tablenames) {
  write_xlsx(tablenames, path = "../tmp/_ROutput_.xlsx", col_names = TRUE, format_headers = FALSE)
}

saveChart <- function(name,  pxwidth = 1024, pxheight = 768, dpi=150) {
  ggsave(filename = paste("../tmp/",name,".png",sep = ""),dpi=dpi, units="in", width=pxwidth/dpi, height=pxheight/dpi)
}

done <- function() {
  file.create("../tmp/done")
  closeAllConnections()
}
