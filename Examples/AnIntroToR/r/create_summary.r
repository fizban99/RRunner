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

# Read the input files
diamonds<- getTable("diamonds")


# Now, let's create our summary (Ctrl-Shift-M for the pipe %>% in RStudio)
Summary <- diamonds %>% 
  group_by(cut, color, clarity, carat2) %>% 
  summarise_all(list(mean)) %>% 
  arrange(desc(cubic))
  
# Write the result
writeResult(tablenames = list("result"=Summary))

# Signal the end of the process
done()

# free up all variables
rm(list=ls())