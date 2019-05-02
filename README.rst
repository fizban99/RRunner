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
* The R Console is used transparently to launch the R scripts, so all output messages are displayed on it. Some antivirus prevent running external files from within Excel, by pre-launching the R Console manually, we can avoid a false positive alert.
* The ranges are sent and retrieved to R through temporary Excel files that you can use while you develop of your scripts, independently of the original Excel. 
* Generate Static plots in R and import them in Excel 
* Sample programs demonstrating the different functions


Installation
============

1. Make sure you have R installed (https://cran.r-project.org/bin/windows/base/)
2. Make sure you have the libraries readxl and writexl installed. Most likely you will want to install the tidyverse libraries as well. 
3. Launch the R Console in SDI mode (single window). You can add –sdi to the Windows shortcut to force the Console to be launched in SDI mode.
4. Import the RRunner.bas module into your Excel project or copy and paste everything (except the first line) into a new module.
5. Make sure you have checked *Microsoft Scripting Runtime* in your Project References.


VBA library usage
=================


Configuration
+++++++++++++++++++++++
You can leave the default configuration parameters as they are. By default, the R scripts will be searched in the same folder as the Excel file and they will be allowed 10 seconds to execute before assuming timeout. The interface files between Excel and R will be called _Input_.xlsx and _Output_.xlsx and will be created in the same folder.
If you modify WORKING_PATH, it should be an absolute path without the ending \. The only exception is ".", which is considered the current folder.


.. code-block:: VB

   ' ###################################################################
   ' Configuration Parameters
   ' ###################################################################
   ' Path to the R Scripts and where the temporary files will be created
   Private Const WORKING_PATH = "."
   ' Time to wait for the R Script answer in milliseconds
   Private Const TimeOutMilliseconds = 10000
   Private Const INTERFACE_IN_FILE_NAME = "_Input_"
   Private Const INTERFACE_OUT_FILE_NAME = "_Output_"
   


RunR2Range
+++++++++++++++++++++++


 


RunR2Plot
+++++++++++++++++++++++




RunRScript
+++++++++++++++++++++++++++++




R library usage
=================

Initialization
+++++++++++++++++++++++++++++


getTable
+++++++++++++++++++++++++++++



writeResult 
+++++++++++++++++++++++++++++


done
+++++++++++++++++++++++++++++