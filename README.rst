RRunner Simple way of running R scripts from Excel VBA
#######################################################

This VBA module for Excel allows running R Scripts from Excel. You can send Ranges to your scripts and retrieve plots from them.

This is a much simpler solution than RExcel. It is simpler to install and, although it has less features, it provides enough functionality to leverage the power of R when you need Excel to perform specific tasks better suited for R.


   .. image:: ./images/ssd1306spi_sm.jpg
      :width: 100%
      :align: center
      
.. contents::

.. section-numbering::


Main features
=============

* Only the default R installation is required. No additional components required to interact with R, although RStudio is recommended for the development of your scripts.
* The R Console is used transparently to launch the R scripts, so all output messages are displayed on it. 
* The ranges are sent and retrieved to R through temporary Excel files that you can use while you develop of your scripts, independently of the original Excel. 
* Static plots 
* Text
* Sample programs demonstrating the different functions


Installation
============

1. 


Library usage
=============


Configuration
+++++++++++++++++++++++



.. code-block:: VBA
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






