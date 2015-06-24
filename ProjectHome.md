Manipulating data in excel is often done faster using the Excel C-API.  We consider this in more detail in this project.


# Introduction #

**_c-api-tools_** is a platfom for building customised data manipulation tools that interface with Excel, at high speed _(native C code)_, using the Microsoft Excel SDK and Excel C-API.

These tools are compiled into an xll which, once loaded, integrate seamlessly with Excel, expanding the number of functions that are on offer within Excel's function suite.

# Distribution #

The latest xll and source code can be download from **[here](http://code.google.com/p/c-api-tools/downloads/list)**

# Details #

Once the xll _Add-in_ has been compiled and **[loaded](http://code.google.com/p/c-api-tools/wiki/HELP_load_configure_addIn)** into Excel, an Add-In function can be called from within Excel to perform the required task.

The example below illustrates how a _c-api-tools_ sorting function **[sort\_ex()](http://code.google.com/p/c-api-tools/wiki/sort_ex)** can be called from within excel to sort the first column in a two column table of numeric data, sorting from low to high:

http://c-api-tools.googlecode.com/files/helpFunctionCall_4.PNG

# Features and Configuration #

  * Ensure that the **[cTools.xll](http://code.google.com/p/c-api-tools/wiki/cTools)** is **[installed and configured](http://code.google.com/p/c-api-tools/wiki/HELP_load_configure_addIn)** correctly.

  * Following this, the Add-In **function features** can be explored in more detail with the help of Excel's _function gui_. This _gui_ displays, amongst other things, the Add-In's **[function signatures and parameter information](http://code.google.com/p/c-api-tools/wiki/HELP_function_parameters)**.

The cTools.xll Add-In should now be visible as a _customised Add-In_. To check this:

  * click on Excel's _insert function_ button. _See arrow 1._
  * if the _cTools_ Add-In is displayed under the _category_ drop-down menu, then the Add-In loaded without error. _See arrow 2._

http://c-api-tools.googlecode.com/files/helpInsertFunction_1.PNG

# Advanced Features #

  * Once loaded, the xll (add-in) will bind automatically with Excel each time Excel starts.
  * This means that all functions in the xll will be available/visible to any spreadsheet or workbook, whenever excel is used.


# References #
_[Microsoft Excel 97 Developer's Kit, Microsoft Press, 1997](http://books.google.com/books?vid=ISBN1572314982)_
