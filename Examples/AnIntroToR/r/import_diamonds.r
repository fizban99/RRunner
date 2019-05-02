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

# We are going to use the diamonds data set that comes with ggplot2
library(ggplot2)
diamonds<- data.frame(diamonds)

# Let's rename columns 8, 9, and 10. 
names(diamonds)[8]<-"length"
names(diamonds)[9]<-"width"
names(diamonds)[10]<-"depth"
# now we have two columns named depth. Let's rename the first one (column 5) to "depthperc." 
names(diamonds)[5]<-"depthperc"
writeResult(tablenames = list("result"=diamonds))
done()