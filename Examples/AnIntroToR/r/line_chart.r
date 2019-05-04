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

# we will use ggplot2 for our charts
library(ggplot2)

# Read the input files
diamonds<- getTable("color_linechart")

# Define the order of the clarity levels
diamonds$clarity <- factor(diamonds$clarity, levels = c("I1","SI2","SI1","VS2","VS1","VVS2","VVS1","IF"))

# We will create a line for each color and see how the number of diamonds of that color change across clarity categories..
ggplot(diamonds, aes(clarity)) + 
  geom_freqpoly(aes(group = color, colour = color), stat="count") + 
  labs(x="Clarity", y="Number of Diamonds", title="Clarity by Color") 
  
# save the chart
saveChart("color_linechart")

# Signal the end of the process
done()

# free up all variables
rm(list=ls())