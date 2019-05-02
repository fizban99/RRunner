RRunner: Simple way of running R scripts from Excel VBA
#######################################################

This VBA module for Excel allows running R Scripts from Excel. You can send Ranges to your scripts and retrieve plots from them.

This is a much simpler solution than RExcel. It is simpler to install and, although it has less features, it provides enough functionality to leverage the power of R when you need Excel to perform specific tasks better suited for R.


   .. image:: ./images/RRunner.png
      :width: 100%
      :align: center
      
.. contents::

.. section-numbering::


Main features
=============

* Only the default R installation is required. No additional components required to interact with R, although RStudio is recommended for the development of your scripts.
* The R Console is used transparently to launch the R scripts, so all output messages are displayed on it. Some antivirus prevent running external files from within Excel. By pre-launching the R Console manually, we can avoid a false positive alert.
* The ranges are sent and retrieved to R through temporary Excel files that you can use while you develop your scripts, independently of the original Excel. 
* Generate Static plots in R and import them in Excel 
* Sample programs demonstrating the different functions


Installation
============

1. Make sure you have R installed (https://cran.r-project.org/bin/windows/base/)
2. Make sure you have the libraries readxl and writexl installed. Most likely you will want to install the tidyverse libraries as well. 
3. Launch the R Console in SDI mode (single window). You can add –sdi to the Windows shortcut to force the Console to be launched in SDI mode. The console has to be running for RRunner to work. It can be minimized, though.
4. Import the RRunner.bas module into your Excel project or copy and paste everything (except the first line) into a new module.
5. Make sure you have checked *Microsoft Scripting Runtime* in your Project References.


VBA library usage
=================


Configuration
+++++++++++++++++++++++
You can leave the default configuration parameters as they are. By default, the R scripts will be searched in the subfolder "R" of the same folder as the Excel file and they will be allowed 10 seconds to execute before assuming timeout. The interface files between Excel and R will be called _Input_.xlsx and _Output_.xlsx and will be created in the same folder.
If you modify WORKING_PATH, it can be an absolute or relative path without the ending \. If using relative paths, they are relative to the folder where the Excel file is located. By default, the scripts are searched in the .\r folder.


.. code-block:: VB

   ' ###################################################################
   ' Configuration Parameters
   ' ###################################################################
   ' Path to the R Scripts and where the temporary files will be created
   Private Const WORKING_PATH = ".\R"
   ' Time to wait for the R Script answer in milliseconds
   Private Const TimeOutMilliseconds = 10000
   Private Const INTERFACE_IN_FILE_NAME = "_Input_"
   Private Const INTERFACE_OUT_FILE_NAME = "_Output_"
   


RunR2Range
+++++++++++++++++++++++

.. code-block:: VB
   
   RunR2Range(script As String, outRange As Range, ParamArray Ranges() As Variant) As Boolean

This function accepts the name of the script (just the name, including the extension), a range where the result will be placed (just the top-left corner cell needs to be indicated) and a set of name-ranges pairs.

E.g.

.. code-block:: VB

   Set Range1 = ActiveWorkbook.Worksheets("Data1").Range("A:C")
   Set Range2 = ActiveWorkbook.Worksheets("Data2").Range("A:B")
   If RRunner.RunR2Range("SampleJoin.r", Range("calculated_values"), "table1", Range1, "table2", Range2) Then
       MsgBox "Done"
   End If
 
This will generate an _Input_.xlsx file with two sheets (they will be called "table1" and "table2" respectively) that will contain the data of Range1 and Range2 and SampleJoin.r script will be called. The script should output the result in a file called _Output_.xlsx in a sheet called "result" (this is hardcoded in the module). The data will be then read from this sheet and placed starting at the top-left corner of the named range called "calculated_values".
Note that although the ranges are referred as the whole columns, only the rows up to the used range will be sent to R. 


RunR2Plot
+++++++++++++++++++++++

.. code-block:: VB
   
   RunR2Plot(script As String, inpRange As Range, outChart As ChartObject, PlotName As String) As Boolean

This function accepts the name of the script (just the name, including the extension), a range from which to read the data and a ChartObject in which to insert the generated chart image. To insert an empty ChartObject, just click on any empty cell, and go to Insert and select any chart type. You can then give a name to this area selecting it and using the usual Name Box (the input box directly to the left of the formula bar).

.. code-block:: VB

   Set ws = ActiveWorkbook.Worksheets("Data1")
   RRunner.RunR2Plot "SampleChart.r", Range("MyPlotData"), ws.ChartObjects("MyChart"), "mychart"

This will generate an _Input_xlsx file with one sheet containing the data in the Named Range "MyPlotData". The sheet will be called "mychart". The R script "SampleChart.r" will be called and it is expected to generate a png file called mychart.png, which will be inserted in the Chart Object "MyChart" after removing any existing image.

RunRScript
+++++++++++++++++++++++++++++

.. code-block:: VB
   
   RunRScript(InputRange As Dictionary, OutputRange As Dictionary, OutputPictures As Dictionary, script As String) As Boolean

This function is the generalisation of the other two. The input ranges are sent as dictionaries using the name as key and the value the actual range.


R library usage
=================

Initialization
+++++++++++++++++++++++++++++

.. code-block:: R
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


getTable
+++++++++++++++++++++++++++++



writeResult 
+++++++++++++++++++++++++++++


done
+++++++++++++++++++++++++++++