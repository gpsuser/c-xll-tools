# c-api-tools : 
[a simple sorting engine for excel using the excel c - api compiled using VSTO]


## Build tools that interface with Excel, using C 

Manipulating data in excel is often done faster using the Excel C-API. We consider this in more detail in this project.

## Introduction

c-api-tools is a platfom for building customised data manipulation tools that interface with Excel, at high speed (native C code), using the Microsoft Excel SDK and Excel C-API.

These tools are compiled into an xll which, once loaded, integrate seamlessly with Excel, expanding the number of functions that are on offer within Excel's function suite.

## Distribution

The latest xll and source code can be download from [here](https://github.com/gpsuser/c-api-tools/archive/master.zip).

## Details

Once the xll Add-in has been compiled and loaded into Excel, an Add-In function can be called from within Excel to perform the required task.

This projct considers how a c based sorting function _sort_ex()_ can be called from within excel to sort the first column in a two column table of numeric data, sorting from low to high:

## Configuration

Ensure that the cTools.xll is installed and configured correctly:
+ in Excel:  File > Options > Addins > Manage Excel Addins > Go > Browse (to the .xll) 

Following this, the Add-In function features can be explored in more detail with the help of Excel's function gui. This gui displays, amongst other things, the Add-In's function signatures and parameter information.

The cTools.xll Add-In should now be visible as a customised Add-In. To check this:

1. Click on Excel's insert function button. 
2. If the cTools Add-In is displayed under the category drop-down menu, then the Add-In loaded without error. 

## Features

Once loaded, the xll (add-in) will bind automatically with Excel each time Excel starts.
This means that all functions in the xll will be available/visible to any spreadsheet or workbook, whenever excel is used.

## References

Microsoft Excel  Developer's Kit, Microsoft Press
