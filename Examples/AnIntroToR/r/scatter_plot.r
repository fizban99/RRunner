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
diamonds<- getTable("price_scatterplot")


# Let's say we want to create a chart that shows how many diamonds of each size (small/medium/large) are in our data.
ggplot(diamonds, aes(carat, price, color=clarity)) + 
  geom_point() + 
  labs(x="Carat Weight", y="Price", title="Price by Carat Weight")
  
# save the chart
saveChart("price_scatterplot")

# Signal the end of the process
done()

# free up all variables
rm(list=ls())