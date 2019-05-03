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

#  let's round the carat values to the nearest 0.25 carat so that our numbers are not all over the place. 
diamonds$carat2 <- round(diamonds$carat/.25)*.25

# Now, let's create our summary. 
Summary <- aggregate(cbind(depthperc, table, price, length, width, depth, cubic)~cut+color+clarity+carat2, data=diamonds, mean)
  
# Write the result
writeResult(tablenames = list("result"=Summary))

# Signal the end of the process
done()

# free up all variables
rm(list=ls())