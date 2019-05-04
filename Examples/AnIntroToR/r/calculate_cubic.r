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

# Read the input files
diamonds<- getTable("diamonds")

# Multiply length, width, and depth together in a new column called "cubic"
diamonds <- mutate(diamonds, cubic=length*width*depth)

# let's round the carat values to the nearest 0.25 carat so that our numbers are not all over the place. 
# this is an alternative way to mutate
diamonds$carat2 <- round(diamonds$carat/.25)*.25

# Select only the calculated column to return less information
diamonds <- select(diamonds, carat2, cubic)

# Write the result
writeResult(tablenames = list("result"=diamonds))

# Signal the end of the process
done()

# free up all variables
#rm(list=ls())