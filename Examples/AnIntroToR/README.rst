An introduction to R for Microsoft Excel users with RRunner
###########################################################

This example sheet is based on the article published by Tony Ojeda (https://districtdatalabs.silvrback.com/intro-to-r-for-microsoft-excel-users)


   .. image:: ./images/an_intro_to_r.gif
      :width: 100%
      :align: center

The scripts assume you have ggplot2, dplyr and reshape2 installed. Remember to open the R Console before clicking on the buttons, since the R scripts are always launched through it. 
The R scripts need to be in the "R" subfolder of the folder where the Excel file is saved. As try the example, the status bar (bottom left corner) will show the part process being performed. Check also the R Console for any errors.

The example shows the usage of some basic but powerful functions of R that can simplify some of your Excel work. This version favours the libraries dplyr and ggplot2 over the basic R functions, since they provide a more readable syntax, so some of the r code will differ from the ones used in Tony's article.

The Excel imports a dataset from ggplot2 and performs some data manipulation, including a summary, a pivot table and a join (vlookup). The last sheet shows some of the R charting capabilities.

