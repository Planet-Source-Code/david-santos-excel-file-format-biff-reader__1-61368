BEEF Reader (Bad Excuse for an Excel File Reader)
-------------------------------------------------

#include <std_disclaimer.h>

Intro
=====

This is the Excel (BIFF) File format parser. Right now it only supports BIFF8/8x (Excel 10/11)


6/10/2005
=========

As of now, BEEF appears to parse the Workbook entry cleanly...
Some remaining Record IDs need to be mapped, and WorkSheet/Cell structure
needs to be implemented to access the data.

6/16/2005
=========

A simple class-based interface has been implemented for BEEF.
This allows you to access the stucture in dot notation.

Fig 1.1
---------

 +------------+
 | BIFFReader |
 +---+--------+
     |
     |   +------------------+
     +---+ cWorkBook Object |
         +---+--------------+   
             |
             |   +-------------------+ 
             +---+ cWorkSheet Object |
                 +---+---------------+
                     |
                     +- Cell Array
                     +- Format Array


Worksheet is implemented as a collection (WorkBook.WorkSheets) and is indexable by the 
worksheet name. For example, if you load an Excel file with 3 worksheets, Sheet1, Sheet2, 
and Sheet3, you will have:


 Dim myReader as BIFFReader

    myReader.OpenBIFF "myfile.xls"

    For each mWorkSheet in myReader.WorkBooks.WorkSheets
	debug.print mWorkSheet.Name 
    Next

Cells are accessed as WorkSheet.Cell(row, "col") as in Cell(1,"A") or numerically, base 1 
indexed as in Cell(1,1).

This is a temporary implementation. The idea implementation would be a Cell Class, though
I am worried about performance issues.





