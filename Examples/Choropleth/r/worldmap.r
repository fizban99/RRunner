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
library(dplyr)
library(scales)
library(ggalt)

# Read the input files
world_data<- getTable("world_population")

# Load the dataframe for the map
world_map <- map_data("world")
# Left join the region field of the maps with the country field of our data
# to get all the countries (even the NA ones)
population.map <- left_join(world_map, world_data, by = c("region"="Country"))
ggplot(population.map, aes(long, lat, group = group))+
  geom_polygon(aes(fill = Population ), color = "#EDF1F9")+
  xlab("") + ylab("") + 
  scale_fill_gradient(low = "#D8E1F2", high = "#315798", na.value="#E0E0E0", name = "Population", labels = comma) +
  coord_proj("+proj=robin +lon_0=0 +x_0=0 +y_0=0 +ellps=WGS84 +datum=WGS84 +units=m +no_defs") +
  theme(panel.background = element_blank(),
      plot.title = element_text(face = "bold"),
      axis.title.x=element_blank(),
      axis.text.x=element_blank(),
      axis.ticks.x=element_blank(),
      axis.title.y=element_blank(),
      axis.text.y=element_blank(),
      axis.ticks.y=element_blank())
  
# save the chart
saveChart("world_population",pxwidth = 800, pxheight = 400, dpi = 75)

# Signal the end of the process
done()

# free up all variables
rm(list=ls())