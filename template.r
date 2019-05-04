# Check if RStudio is running to set the working directory to the script directory
# https://stackoverflow.com/questions/35986037/detect-if-an-r-session-is-run-in-rstudio-at-startup
is.na(Sys.getenv("RSTUDIO", unset = NA))
if (!is.na(Sys.getenv("RSTUDIO", unset = NA))) {
  # Get current directory
  current_dir <- dirname(rstudioapi::getSourceEditorContext()$path)
  # Set working directory to current directory (script directory)
  setwd(current_dir)
} else {
  # If sourced https://stackoverflow.com/questions/13672720/r-command-for-setting-working-directory-to-source-file-location-in-rstudio
  this.dir <- dirname(parent.frame(2)$ofile)
  setwd(this.dir)
}
# Include the excel helper functions
source("excelhelper.r")

# libraries for data manipulation
library(dplyr)
library(reshape2)

###########################
# Custom Code starts here #
###########################

# Read the input files
table1 <- getTable("table1")

# Process the data
res <- table1

# Save any plots
saveChart("mychart")

# Write the result
writeResult(list("result" = res))

# Signal the end of the process
done()
# free up all variables. You might want to comment this out while testing your script in RStudio
rm(list=ls())