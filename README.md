# A fork of Ben's POIUtility, with a shiny new exterior and motor

I've brought POIUtility and custom tags up to date to use the latest Apache POI 4.0.1, make it not dependent on the server libraries, support both XLS and XLSX, and some new features and a revamped CSSRule.cfc.  The POIUtility.cfc has been updated to use javaLoader and now uses the CSSRule for style management.  Both XLS and XLSX can be read or written.  Additional parameter support for data row start, number of rows to read, column start, and number of columns to read have been added.  It also does a hard check beyond the get Last Row and Last cell methods, scanning the cells for non blank.  This prevents empty rows being read in the query.  
 
 The query mechanism now returns the query with the proper column names if a header row is available.

 
* Added update boolean to poi:cell which will try to retrieve existing cell and only update the value, preserving cell formatting if   template is used.(4/24/2019)
* Fixed poi:document EvaluateFormulas to evaluate all cells in xlsx (tested with single sheet formula references ) (4/24/2019)
* Added comment and comment author support for cells. (4/12/2019)
* Added support for Apache POI 4.0.1
* Added  Mark Mandel's JavaLoader.cfc
* hex color find nearest for xls
* Now Supports xlsx ( Just add an createXLSX="true" attribute in the document tag )
* Revamped CSSRule.cfc to cfscript, added generic border style and color settings, added rowSpan, when used with colSpan, will merge regions.

### Here is a color pattern test showing generation of both XLS and XLST
<img src="https://github.com/wiggick/POIUtility.cfc/blob/master/docs/ugly_spreadsheet.PNG">

### Comments now supported
<img src="https://github.com/wiggick/POIUtility.cfc/blob/master/docs/comment_support.PNG">

# POIUtility.cfc


The POIUtility.cfc is a ColdFusion component that helps you read Microsoft Excel files into ColdFusion 
queries as well as convert ColdFusion queries into multi-sheet Microsoft Excel files.

## Features

* Supports creation of both xls ans xlsx 
* Optional Header Row.
* Optional Header Start Row
* Optional Data Start Row
* Optional Column Start Row
* Optional number of columns to read
* Optional number of rows to read
* Basic CSS style definitions for header, row, and alternating row using a CSSRule.cfc engine
* Write single or multiple Excel sheets at one time.
* Read entire workbook into array of sheets or read in single sheet.

# POI ColdFusion Custom Tag Features

* Create Excel documents in either xls or xlsx using ColdFusion custom tags.
* Write Excel file to file server or to a ColdFusion variable (or both).
* CSS control at the global, column, row, and cell levels (with proper cascading).
* Date and Number mask support.

These ColdFusion custom tags allow you to create native Microsoft Excel binary files. They create PRE-2007 compatible files. The following is a list of the currently supported tags and the current attributes.

__NOTE__: All tags in the POI systems require the use of both an OPENING and CLOSING tag. If you leave out a closing tag (or self-closing tag), you will get unexpected results.


## Document

Name: [optional] If provided, will store a copy of the Excel file in a ByteArrayOutputStream that can easily be converted to a byte array and streamed to the browser using CFContent or written to the file system using CFFile.

File: [optional] If provided, will store a copy of the Excel file at the given expanded file path.

Template: [optional] If provided, this will read in and use an existing Excel file as the base for the new file (it does not affect the template, only copies it's data). 

Style: [optional] Sets default CSS styles for all cells in the document.

CreateXLSX: [optional] boolean if true will create xlsx (default)

EvaluateFormulas: [optional] will evaluate formulas after creation/edit.  Useful for working with templates with predefined formulas.

__Note__: Name and File are optional, but ONE of them is required.


## Classes

No functional value other than containership at this time.


## Class

Name: The name of the class (to be used as a struct-key) holding the given CSS styles.

Style: The CSS style for this class.

Note: You can use the class name "@cell" to override the default cell style for the entire workbook.


## Sheets

No functional value other than containership at this time.


## Sheet

Name: The name of the sheet to be displayed in the tab at the bottom of the workbook. 

FreezeRow: [optional] The one-based index of the row you want to freeze.

FreezeColumn: [optional] The one-based index of the column you want to freeze.

Orientation: [optional] The print orientation of the sheet. Can be portrait (default) or Landscape.

Zoom: [optional] The default zoom of the sheet as a percentage (ex. 100%).


## Columns

No functional value other than containership at this time.
This section is optional.


## Column

Index: [optional] The zero-based index of the column. By default, this will start at zero and increment for each column.

Class: [optional] The class names (defined above) that should be applied to this column. This can be a single class or a space-delimited list of classes (to be taken in order).

Style: [optional] The CSS styles that should be applied to this column.

Freeze: [optional] Boolean value to determine if this column should be frozen in the document.

Update: [optional] if set and cell exists in sheet, will only set the value of the cell, ignoring any formatting options. (might change this to still retrieve the cell if it exists, allowing for partial style updates.)


## Row

Update: [optional] When loading in a template, specifying update as true will read in the existing row of a template for cell editing

Index: [optional] The zero-based index of this row. By default, this will start at zero and increment for each row. If you set this manually, all subsequent rows will start after the previous one.

Class: [optional] The class names (defined above) that should be applied to this row. This can be a single class or a space-delimited list of classes (to be taken in order).

Style: [optional] The CSS styles that should be applied to this row.

Freeze: [optional] Boolean value to determine if this row should be frozen in the document.


## Cell

Type: [optional] Type of data in the cell. By default, everything is a string. Currently, can also be Numeric, Date, or Formula.

Index: [optional] The zero-based index of this cell. By default, this will start at zero and increment for each cell. If you set this manually, all subsequent cells in this row will start after the previous one.

Value: [optional] The value to be used for the cell output. If this is not provided, then the GeneratedContent of the cell tag will be used (space between the opening and closing tags).

Comment: [optional] add a comment to a cell

CommentAuthor: [optional] Specify author of comment defaults to "Apache POI"

ColSpan: [optional] Defaults to one; allows you to create merged cells in a horizontal way.

RowSpan: [optional] Defaults to one; allows you to create merged cells in a vertical way.  Used with ColSpan to create merged region.

NumberFormat: [optional] The number mask of the numeric cell. Only a limitted number of masks are available.

DateFormat: [optoinal] The date mask of the date cell. Only a limited number of masks are avilable.

Class: [optional] The class names (defined above) that should be applied to this cell. This can be a single class or a space-delimited list of classes (to be taken in order).

Style: [optional] The CSS styles that should be applied to this cell.

Alias: [optional] Creates a pointer to the given cell for use within another cell formula. When being referenced in a cell formula, use the @ sign (ex. "SUM( @Start:@End )").

## CSS 
Utility tag for accessing CSSRule methods directly.  Usefull for unit testing and iteration of styles and colors

Method: the method to call (limited to specified list 

Var: unused placeholder for future expansion

Result: variable where the result of the method will be returned to.

## Available Number Formatting Masks

* "General"
* "0"
* "0.00"
* "#,##0"
* "#,##0.00"
* "($#,##0_);($#,##0)"
* "($#,##0_);\[Red\]($#,##0)"
* "($#,##0.00);($#,##0.00)"
* "($#,##0.00_);\[Red\]($#,##0.00)"
* "0%"
* "0.00%"
* "0.00E+00"
* "# ?/?"
* "# ??/??"
* "(#,##0_);\[Red\](#,##0)"
* "(#,##0.00_);(#,##0.00)"
* "(#,##0.00_);\[Red\](#,##0.00)"
* "_(*#,##0_);_(*(#,##0);_(* \"-\"_);_(@_)"
* "_($*#,##0_);_($*(#,##0);_($* \"-\"_);_(@_)"
* "_(*#,##0.00_);_(*(#,##0.00);_(*\"-\"??_);_(@_)"
* "_($*#,##0.00_);_($*(#,##0.00);_($*\"-\"??_);_(@_)"
* "##0.0E+0"
* "@" - This is text format.
* "text" - Alias for "@"

## Available Date Formatting Masks

* "m/d/yy"
* "d-mmm-yy"
* "d-mmm"
* "mmm-yy"
* "h:mm AM/PM"
* "h:mm:ss AM/PM"
* "h:mm"
* "h:mm:ss"
* "m/d/yy h:mm"
* "mm:ss"
* "[h]:mm:ss"
* "mm:ss.0"

[1]: http://www.bennadel.com
