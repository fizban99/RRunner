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

# library for data manipulation
library(dplyr)
library(reshape2)

# Read the input files
diamonds<- getTable("diamonds")

# Then, we'll use the dcast function to get our data into the same pivot table format.
# we're taking the color, clarity, and price columns from the diamonds data frame, 
# casting (pivoting) them out by color (rows) and clarity (columns),
# and calculating the average price for each combination. 
pivot_table <- dcast(diamonds[,c('color','clarity','price')], color~clarity, mean)

# Write the result
writeResult(tablenames = list("result"=pivot_table))

# Signal the end of the process
done()

# free up all variables
rm(list=ls())