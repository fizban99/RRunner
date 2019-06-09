An introduction to R for Microsoft Excel users with RRunner
###########################################################

This example sheet shows how to create a choropleth map (thematic map in which areas are shaded or patterned in proportion to the measurement of the variable being displayed on the map) in Excel using RRunner


   .. image:: ./images/choropleth.png
      :width: 100%
      :align: center

The scripts assume you have ggplot2, dplyr, scales and ggalt installed. If not, go on and install them with RStudio, for example. Remember to open the R Console before clicking on the buttons, since the R scripts are always launched through it. 
The R scripts need to be in the "R" subfolder of the folder where the Excel file is saved. As you try the example, the status bar (bottom left corner) will show the part process being performed. Check also the R Console for any errors.

The example uses a Robinson projection, which is not the default that ggplot2 uses, but it has a more "modern" look. The colours have been also tuned to be more pleasant, but you can change them directly in the script.