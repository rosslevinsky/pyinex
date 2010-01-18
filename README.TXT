Pyinex - a project to embed Python in Excel
README.TXT
-------------------------------------------

Version 0.83 alpha  - January 18, 2010
Copyright (c) 2010, Ross Levinsky 


Introduction
------------

Pyinex is a project to embed the Python interpreter in Excel. This is in contrast to the usual approach to linking the two systems, in which Python scripts run in an out-of-process COM server and Excel makes a cross-process call to use Python functionality. In Pyinex, Python runs from a DLL in the Excel process space.

This embedding offers the advantage of speed (no cross-process calls) and reduced fiddliness. There is no setting up of COM servers and their associated security - you simply write a Python script in a text file, and call it from within Excel.

This opens up the entire world of Python libraries to Excel users, and allows reduced dependence on VBA code, which is always hidden inside the Excel file. With program logic in separate, easily-viewable Python scripts, a new level of testing, visibility, and control of spreadsheet logic is possible. Excel can serve as a user scratchpad, UI, or graphing tool, and leave the "serious" computation to Python code.


Requirements
------------

- Excel 2002 or 2007. I have only tested it against 2002 and 2007, as I don't have a copy of 2003, but it is very likely to work there as well.

- Python 2.5, 2.6, or 3.1 (the python.org distribution) installed on your machine. Other distributions may work, but Pyinex hasn't been tested against anything else.

Pyinex is developed with Visual C++ 2008 Express Edition, available free of charge from Microsoft. You can download it from their website.


Basic operation
---------------

Pyinex is an Excel extension library - an XLL - written in C++, using the open-source XLW library. It currently provides six functions to Excel:

1) PyCall( filename, 
   	   function, 
           15 more arguments to pass to the function )

PyCall loads up the Python script specified in the filename parameter (technically, it imports the module) and calls the function specified in the function parameter, passing as many of the 15 additional arguments as that function signature requires. It returns one of three things: a single value, a single row, or a two-dimensional array (possibly with a ragged right edge) of data. 

It's important to note the distinction between a single row and a single vertical column of returned data. Pyinex takes any tuple or list containing only fundamental types (int, float, string, bool) and returns it to Excel as a single horizontal row. For example, a returned Python variable of [1,2,"me","you"] or (1,2,"me","you") is rendered as a single row in Excel, showing "1,2,me,you" in successive horizontal cells.

A vertical row is treated as a single-column two-dimensional matrix (N rows by one column). A Python return value of ((1,), (2,), ("me",), ("you",)) is returned to Excel as a vertical strip of cells.

This all applies in reverse, too - a horizontal strip of Excel cells (as an input argument) turns into a tuple containing fundamental data types, and a vertical strip is mapped to a nested set of tuples.

Repeating the previous example, but in the other direction: passing in the horizontal strip of cells "1,2,me,you" as an argument produces (1,2,"me","you") in Python, and passing the same input vertically produces ((1,), (2,), ("me",), ("you",)) in Python.

In summary, Pyinex treats single rows as vectors, not matrices, and single columns as matrices, not vectors.

2) PyConsole( required showConsole flag (TRUE or FALSE),
   	      optional x position of the upper-left corner (pixels),
	      optional y position of the upper-left corner (pixels),
	      optional width (pixels),
	      optional height (pixels),
	      optional preserveLocation flag (TRUE or FALSE) )

PyConsole() toggles the visibility of a Windows console (the black screen one sees in console-mode applications, such as the "cmd" shell). When the Pyinex XLL first loads into Excel, it redirects the stdout and stderr streams to that console, which means that print statements in Python scripts, Python interpreter error messages, and Pyinex error messages all appear in that window.

You may optionally specify the position of the upper-left corner of the console, in pixels. You must specify both x and y positions if you specify one, and they must be greater than zero and less than the size of the display. Pyinex will flag any errors in this regard.

You may optionally specify the width and height of the console, in pixels. You must specify both if you specify one, and their size must be greater than zero.

The optional preserveLocation flag determines whether or not the console location is reset to the specified location/size when it is cycled from visible to hidden and back to visible. The default for the flag is TRUE, which means that a visibility cycle doesn't reset the location/size. If you wish to programmatically reset location/size, set this to FALSE and toggle the showConsole flag.

3) PyClearConsole( TRUE or FALSE )

This function clears the console when the input is toggled to TRUE. If you leave it set to TRUE, console output will clear with every calculation cycle; depending on the calculation order Excel chooses, this may leave you with full output from your current calculation cycle, or a fraction of the results, or no results at all. So, don't leave it set to TRUE - just toggle it when you need to clear the screen.

4) PyVerbose( integer )

This sets the verbosity of the Python interpreter error messages (equivalent to setting the -v flag when calling python from the command line). Values must be greater than zero, and in the current interpreter code, the maximum value of interest is approximately three.

5) PyLoadedLibrary( "python" | "pyinex" )

This displays the current loaded Python and Pyinex DLL names, depending on which name is passed in (comparison is case-insensitive). It is useful to validate that you're actually using the XLL and python version that you think you are using (a mistake that's easy to make if you have multiple versions of the add-in installed in Excel).

6) PyModuleFreshnessCheck( optional TRUE or FALSE )

When passed TRUE, each execution of PyCall polls the last write time of the passed-in module file. If the file has changed since last visited, it is reloaded. This eases development, because one doesn't have to restart Excel to pick up changes to modules, but slows execution, because of the constant file polling. The default behavior is to poll (internal setting of TRUE).

Passing FALSE turns off this polling - the first version of a module that's loaded will be used for the lifetime of the Excel process (or until TRUE is passed to this function). This is useful for speeding production spreadsheets in which the code is not expected to change.

Passing nothing causes the function to simply return the current value of the setting. 


Python extensions
-----------------

Pyinex provides six functions that extend Python. These live in the module "pyinex", which is automatically loaded into the Python interpreter at startup. You do not need to call "import pyinex", though you may do so if you wish to alias the module name ("import pyinex as youraliashere").

1) CallerA1() - provides the name of the calling Excel cell in A1 format

2) CallerA1Full() - provides the name of the calling Excel cell in A1 format, with the full sheet name prepended. Example: [PythonExtensionTest.xls]Sheet1!C12

3) CallerR1C1() - provides the name of the calling Excel cell in R1C1 format

4) CallerR1C1Full() - provides the name of the calling Excel cell in R1C1 format, with the full sheet name prepended. Example: [PythonExtensionTest.xls]Sheet1!R12C3

5) CallerSheet() - provides the name of the sheet on which the calling cell resides. Example: [PythonExtensionTest.xls]Sheet1

6) Break( boolean clearBreak ) - queries Excel to see if the user has pressed the Esc key during calculation, and returns True if so. This is useful to check in the midst of any long-running calculation in Python, with the usual desired behavior being to abort the running calculation and return control to Excel. 

The optional clearBreak boolean (default is False) tells Excel to clear the Esc signal for any future queries of this function during a single calculation cycle. If set to True, the break request is cleared, and any other cells that query Break() will receive a False until the user presses Esc again. The behavior you'll most often want is to NOT clear the Esc request (hence the False default); this allows you to stop all calculations that query Break() with a single press of Esc.


Examples
--------

There are five examples provided, each of which comprises an identically named .xls/.py file pair. You'll need to edit the cell in the .xls file that points to the .py file (cell C1 of the first worksheet of each workbook); it must contain the location of the .py file on your local install.

1) PyinexTest - tests a variety of Pyinex functions. It shows how various Python data types are returned; the use of array functions; how to do proper string-based lexicographic sorting in all supported versions of Python (see the stringize() function in the .py file); and how to retrieve data from the web.

2) PythonExtensionTest - shows the use of the Python extension caller-name retrieval functions.

3) PythonBreakTest - shows the use of the Python extension Break() function.

4) NumPyDemo - displays a trivial use of the NumPy library. This only works on Python 2.5 and 2.6, as a 3.1 version of NumPy has not been released. You'll need to have NumPy installed on your system (available from numpy.scipy.org).

5) UnicodeTest - demonstrates how Pyinex handles strings taken from Excel. Behavior varies depending on the combination of Python and Excel in use (see "String handling," below).
 

Module importing and reloading
------------------------------

Because there is no obvious "main" script in an Excel file (i.e., you can specify different script files in each of multiple calls to PyCall on a single spreadsheet), Pyinex treats each Python file as a module to be imported. This means that any code outside of function definitions will be run at module import time, exactly as is the case when the standalone Python interpreter imports a module.

During development, Python scripts are typically edited while Excel remains running. It's very inefficient to require quitting and restarting of Excel simply to reload the changed script, so one of two things must be done: allow for manual specification of reloading, or automatically reload when the script file changes. As a convenience, Pyinex implements the latter approach (though this can be toggled with PyModuleFreshnessCheck).

The side effect of this choice is obvious: if you have global variables set in your running module and you edit and save the module file, Pyinex will reload the module, thereby wiping out the variables. This is largely a problem for development (as modules generally aren't edited during production runs).


String handling
---------------

Pyinex needs to be aware of how both Excel and Python handle strings. The situation is moderately complex, with growing sophistication as both Excel and Python evolved.

The XLL interface in Excel 2002/2003 can only handle strings in which each character uses a single byte, the most common example of which is standard ASCII strings. These versions of Excel can open a workbook with Unicode strings (perhaps put there by Excel 2007 or some other application), but can't pass those strings as Unicode to the XLL; they're converted to ASCII to the extent possible, and all extended characters are lost (typically replaced by ASCII '?' characters).

Excel 2007 natively passes and receives wide character strings, which I am nearly certain are UTF-16 encoded Unicode characters. The documentation isn't precise on this point.

Python 2.x treats single-byte and Unicode strings as separate types, with a whole suite of conversion routines available to translate Unicode into various encodings.

Python 3.x eliminates the separate Unicode type and treats all strings as Unicode, with the caveat that the former "Unicode" type was renamed to "str". It's probably clearer to say that it eliminated the Python 2.x string type, renamed the Python 2.x "Unicode" type to "str", and created a new "bytes" type that can represent single-byte character strings.

We thus have four specific combinations of Excel/Python versions to consider:

1) Excel 2002/2003 and Python 2.x - Excel passes and receives only single-byte-width characters. Any Unicode strings in the workbook are downconverted to ASCII and passed to Python as string objects, with all extended characters (value > 0x7F) lost. 

Strings coming back from Python (which by definition are single-byte-width) to Excel are properly converted. Unicode objects are downconverted and returned as single-byte ASCII strings if they contain only ASCII characters; otherwise, an error message is returned to the Excel cell.

2) Excel 2002/2003 and Python 3.x - Excel passes and receives only single-byte-width characters. Any Unicode strings in the workbook are downconverted to ASCII with all extended characters lost. These ASCII strings are then passed to Python and upconverted to Python 3.x string objects (which are essentially Unicode strings). This seems perverse, but the XLL interface doesn't handle Unicode.

Strings coming back from Python, where they are by definition Unicode, are converted to single-byte-width ASCII if possible. If the conversion fails, an error message is returned to the Excel cell.

3) Excel 2007 and Python 2.x - Excel natively passes and receives Unicode strings, but Python treats single-byte-width strings and Unicode strings differently.

Because the vast majority of Python programs (and programmers) aren't Unicode-aware, Pyinex tries to pass string objects into Python whenever possible. All Unicode strings coming from Excel are examined to see if they consist of only ASCII characters, and if so, a string object is passed in. If not, a Unicode object is created.

All strings coming back from Python are converted to Unicode strings in Excel.

4) Excel 2007 and Python 3.x - both Python and Excel are Unicode-aware, so they treat all strings as Unicode. This is the simplest/best case; as older versions of Excel and Python recede into history, the complexities of string handling will disappear.


Limitations
-----------

- The Excel interface is fundamentally limited - essentially, all data can be at most a two-dimensional region, and one has to use clever/fiddly formatting within any such region in order to simulate any more-complicated data structure (e.g., XLW's ArgumentList class). Because this is a non-robust approach, I have shied away from supporting it, preferring to stick to a simple set of allowable input and output styles.

- In particular, Pyinex doesn't support the return of dictionaries, largely because of the complexity of examining all of the keys and values for their ability to be "flattened" into a shape that Excel can reasonably render. This is not a terrible limitation, because the Python script can do the flattening and return the info as a tuple, but I may relax this prohibition in the future, permitting well-formed dictionaries to be returned.

- Excel internally treats all numbers on a sheet as floats, even if they're formatted as integers. Consequently, when they're passed in to Python, they are converted to floats. This is problematic when using a passed-in Excel number as, say, a list index, because those must always be Python integers or longs. The solution is simple - explicitly convert from float to int in the Python script (e.g., excelNum = int(excelNum)).

- Pyinex has been tested on Excel 2002 (from the Office XP suite) and Excel 2007. I haven't tested it on Excel 2003 simply because I don't own a copy.

- PyCall is single-threaded. Excel 2007 supports multiple calculation threads in Excel and one can mark XLL functions as thread-safe, so this is theoretically possible, but one first has to contend with the details of using the Python interpreter in a thread-safe fashion. I have not had a chance to do this research.

- All calls go to a single instance of the Python interpreter, despite the theoretical ability to embed multiple interpreters in a single process. This also awaits further research.

- There is no evident way (yet) to interrupt the Python interpreter in mid-calculation, so it's possible to hang Excel with a badly written Python script. This is likely fixable if we produce a modified Python interpreter; it could be programmed to periodically look for Excel interrupts. It may also be fixable with a yet-to-be-defined yield discipline for scripts.

- PyCall is limited to 15 function arguments. Excel doesn't formally permit the use of a variable number of arguments in calls to an XLL function; one has to specify the fixed number of arguments that the XLL function can accept, and if less are passed in by the user, Excel makes up the difference by tacking on empty arguments to the user's list and passing the resulting agglomeration to the XLL function. Pyinex code is written to take the entire argument list (all arguments actually specified by the user, plus the empty arguments that Excel adds on) and prune it down to the specific number of arguments that the Python function requires. Due to various arcane technical limitations of Excel 2002/2003 and XLW, the maximum argument count is 15. I doubt this will be a serious limitation, as each argument can be a two-dimensional array of data, and there are many ways to store intermediate states of computation in global Python objects.

- PyCall doesn't work properly with scripts that run their own GUI message loop. In particular, Python's matplotlib plotting library doesn't work correctly when it is rendering to interactive screens, because those screens depend on one of a number of widget toolkit message loops (i.e., Tk, Wx, or Qt message loops). This may be possible to fix - I believe that the Python shell IPython handles this by placing those GUI message loops into separate threads - but I haven't had a chance to experiment with ideas like this in Pyinex.


Interesting technical points
----------------------------

1) The dynamic repointing of CRT output streams in Utils/LoadedCRT.cpp

2) The loading and setup of Python extension modules and the experimental insertion of the module name into the global namespace in Pyinex/PythonExtension.cpp.

3) The conversion of Excel column numbers to column names in Pyinex/PythonExtension.cpp, and the test code that validates the algorithm in Pyinex/ExcelColnameValidation.py.


Future development
------------------

Pyinex is a work in progress; some of the development ideas I am considering include:

- Named arguments to functions. This is hard(er), mostly because there's no obviously good UI paradigm for this in Excel. The XLW team has their ArgumentList class to try to address this problem, but it's clunky.

- Evaluation of arbitrary Python code (submitted as text from Excel, instead of in a standalone file). This is likely easy.

- Programmatic control over reimportation of modules. This now happens automatically when a module has been modified, but it's straightforward to implement a "reload" command.

- Programmatic control over reinitialization of the Python interpreter. This can be done by quitting and restarting Excel, but it should be allowed programmatically.

- Make PyCall thread safe, and possibly allow the use of multiple interpreters. This is only useful for Excel 2007 and its successors.

- Replace polling for module updates with "something better." Windows file notification doesn't work reliably with network drives, so the most likely candidate is a background thread that polls at some relatively low frequency.

- Improve text output and error logging. Output and errors should be capturable to text files, and error  visibility should be controlled by a programmatically-settable logging verbosity switch (distinct from the Python interpreter's verbosity switch).

- Wrap PyObject references in a smart pointer class. This would greatly simplify the reference counting in the code. I assume there's an open-source class that handles this well, so it may only be a matter of research.

- A test suite. The Excel<->Python variable conversion code is ugly, and it needs a test suite (beyond my personal scripts) to validate it.


Obtaining the latest version
----------------------------

Pyinex source code is available at code.google.com/p/pyinex, and will generally include a binary distribution.


XLW
---

Pyinex is written using the open-source XLW library (available at xlw.sourceforge.net), which greatly facilitates the production of XLLs. With the exception of one (presumed) bug fix to XLW 4.0's code, it is used as supplied by the authors. I have submitted the proposed fix to them for inclusion in their future releases.


License
-------

Pyinex is released under the modified BSD license; the license does not require the advertising clause present in the original BSD license. In particular, you can use Pyinex freely for commercial projects and products, and you do not need to expose or otherwise redistribute the resulting source code. You do need to include elements of the copyright notice, though - see the license comment at the top of all Pyinex files for details.

As a matter of curiosity, I would very much appreciate a brief email (to pyinex at toponia dot com) telling me if you are using it, but you are not required to do so. I will keep all such communications absolutely private.


Potential contributors
----------------------

I welcome your ideas, requests, bug reports, and general assistance. If you are a reasonable C++ programmer and know something about low-level Excel or Python programming, please get in touch with me if you'd like to pitch in. For now, I reserve the right to pass editorial judgment on submitted code (i.e., I am the Benevolent Dictator Of Pyinex, though perhaps not for life).


Contact
-------

Please send any questions, comments, or bug reports to pyinex at toponia dot com.
