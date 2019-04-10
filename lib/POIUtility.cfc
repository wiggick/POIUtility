	/**
	* Handles the reading an writing of Excel files using the POI
	* package that ships with ColdFusion.
	* Special thanks to Ken Auenson and his suggestions:
	*
	* http://ken.auenson.com/examples/func_ExcelToQueryStruct_POI_example.txt
	* Special thanks to (the following) for helping me debug:
	*
	* Close File Input stream
	* 	- Jeremy Knue
	*
	* Null cell values
	* - Richard J Julia
	* - Charles Lewis
	* - John Morgan
	*
	* Not creating queries for empty sheets
	* - Sophek Tounn
	*
	* @author Ben Nadel, Chris Wigginton (rewrite for Apache POI 4.0.1 )
	* @version 4.0
	* @date 4/5/2019
	*
	*  Update History:
	*  04/09/2019 Chris Wigginton
	* 	WOW Really? over 12 years since an update :-)
	*   Read and Write both XLS and XLSX
	*   now uses JavaJoader
	*   now integrated with CSSRule.cfc for style management
	*   support for specifying when the datarow start
	*   support for number of rows to read
	*   supports for column start and number of columns to read
	*   hard check for empty _BLANK rows when reading in spreadsheet
	*
	*  04/04/2007 - Ben Nadel
	*  Fixed several know bugs including:
	*  - Undefined query for empty sheets.
	*  - Handle NULL rows.
	*  - Handle NULL cells.
	*  - Closing file input stream to OS doesn't lock file.
	*
	*  02/01/2007 - Ben Nadel
	*  Added new line support (with text wrap). Also set the sheet's
	*  default column width.
	*  01/21/2007 - Ben Nadel
	*  Added basic CSS support.01/15/2007 - Ben Nadel
	*  Laid the foundations for the CFC.
	**/

/**
* @displayName POIUtility
* @hint Handles the reading and writing of Microsoft Excel files using POI and ColdFusion.
* @accessors false
* @output false
*/
component {


	/**
	* Init
	* @hint Returns an initialized POI Utility instance.
	* @output false
	*/
	public any function Init(){

		VARIABLES.poiPath =  GetDirectoryFromPath ( GetCurrentTemplatePath() ) & "tags/poi/";
		VARIABLES.loadPaths = [];
		VARIABLES.isXLSX = true;
		VARIABLES.loadPaths[1] = replace( "#VARIABLES.poiPath#apache/poi-4-0-1.jar","\","/","all");
		VARIABLES.loadPaths[2] = replace( "#VARIABLES.poiPath#apache/poi-ooxml-4-0-1.jar","\","/","all");
		VARIABLES.loadPaths[3] = replace( "#VARIABLES.poiPath#apache/lib/commons-collections4-4.2.jar","\","/","all");
		VARIABLES.loadPaths[4] = replace( "#VARIABLES.poiPath#apache/xmlbeans-3.1.0/lib/xmlbeans-3.1.0.jar","\","/","all");
		VARIABLES.loadPaths[5] = replace( "#VARIABLES.poiPath#apache/poi-ooxml-schemas-4.0.1.jar","\","/","all");
		VARIABLES.loadPaths[6] = replace( "#VARIABLES.poiPath#apache/lib/commons-compress-1.18.jar","\","/","all");

		VARIABLES.workbookFactoryClass  = "org.apache.poi.ss.usermodel.WorkbookFactory";
		VARIABLES.cellRegionClass       = "org.apache.poi.ss.util.CellRangeAddress";


		//workbookFactoryClass and the CSSRule do all the heavyLifting

		//Place any specific engine XLS or XLSX here that the POIUtility needs to use
		VARIABLES.XLSXClasses = {
			formulaEvaluatorClass = "org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator"
		};

	    VARIABLES.XLSClasses = {
	    	formulaEvaluatorClass = "org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator"
		};
		
		VARIABLES.javaLoader = createObject("component", "lib.tags.poi.javaloader.JavaLoader").init(VARIABLES.loadPaths);

		//Common to both XLS and XLSX
		VARIABLES.WorkBookFactory = VARIABLES.javaLoader.create(  VARIABLES.workbookFactoryClass ).Init();
		return this;

	}



	/**
	* @hint Takes the standardized CSS object and creates an Excel cell style.
	* @output false
	*/
	private any function GetCellStyle(required any WorkBook, required any CSS ){

		// Create a default cell style object.
		LOCAL.Style = ARGUMENTS.WorkBook.CreateCellStyle();
		VARIABLES.CSSRule.ApplyToCellStyle(PropertyMap = ARGUMENTS.CSS, Workbook= ARGUMENTS.WorkBook , CellStyle=LocalStyle);
		return LOCAL.Style;
	}

	/**
	* @hint Create or return appropriate POI engine which would have a correctly initialized CSSRule for the given type
	* @type Determines how the Workbook is created, read is read/close, create only creates the workbook 
	*/
	public any function GetPOIWorkBookObjects( required string FilePath, string type="read", boolean createCSSRule= false ){

		if(! ListFind("read,write", ARGUMENTS.type ) ){
			throw(type="POIUtiiltyException", Message="Invalid type: #ARGUMENTS.type#");
		}

		LOCAL.fileExtension = ListLast(ARGUMENTS.FilePath,".");

		if(! ListFindNoCase("xls,xlsx", LOCAL.fileExtension ) ){
				throw( type="POIUtilityException", message="Invalid extension type #local.fileExtension# for #ARGUMENTS.FilePath#");
		}

		LOCAL.isXLSX = ( lcase(local.FileExtension) eq "xlsx" );

		if( ARGUMENTS.type eq "read" ){
			

			LOCAL.FileInputStream = CreateObject( "java", "java.io.FileInputStream" ).Init(
			JavaCast( "string", ARGUMENTS.FilePath )
			);
		
			LOCAL.WorkBook = VARIABLES.WorkBookFactory.Create( LOCAL.FileInputStream );

			LOCAL.FileInputStream.Close();

		}else{
			//writing we handle the fileOutput stream later
			LOCAL.WorkBook = VARIABLES.WorkBookFactory.Create( JavaCast( "boolean",LOCAL.isXLSX ) );		
		}
		
		// Now that we have crated the Excel file system,
		// and read in the sheet data, we can close the
		// input file stream so that it is not locked.
		//LOCAL.FileInputStream.Close();
		if( ARGUMENTS.createCSSRule ){
			LOCAL.CSSRule = CreateObject( "component", "lib.tags.poi.CSSRule" ).Init( isXLSX =  LOCAL.isXLSX , javaLoader = VARIABLES.javaLoader, WorkBook = LOCAL.WorkBook );
			LOCAL.result =  {WorkBook = LOCAL.Workbook, CSSRule = LOCAL.CssRule };
		}else{
			LOCAL.result = { WorkBook = LOCAL.Workbook };
		}
		
		return LOCAL.result;
	}

	/**
	* @hint Reads an Excel file into an array of strutures that contains the Excel file information OR if a specific sheet index is passed in, only that sheet object is returned
	* @FilePath The expanded file path of the Excel file to be used as the template
	* @HasHeaderRow Flags the Excel files has using the first data row a header column. If so, this column will be excluded from the resultant query.
	* @SheetIndex If passed in, only that sheet object will be returned (not an array of sheet objects)
	* @RowsToRead  If 0, then all, otherwise limit
	* @ColumnsToRead optional indicates number of columns to retrieve, 0 will grab to max columns available
	* @ColumnStart optional  1 based index identifies column start
	* @HeaderRowStart optional (requires true for HasHeaderRow) 1 based index.
	* @DataRowStart optional, 1 based index, row index when to start reading data
	* @output false
	*/
	public any function ReadExcel(required string FilePath ,boolean HasHeaderRow=true,
		numeric SheetIndex  = -1,
		numeric RowsToRead = 0,
		numeric ColumnsToRead = 0,
		numeric ColumnStart,
	    numeric HeaderRowStart,
	    numeric DataRowStart
		 ){

		//some validation here, we expect the user to pass in 1 based index rules
		if( StructKeyExists(ARGUMENTS,"HeaderRowStart") ){
			If( ! ARGUMENTS.HasHeaderRow ){
				throw(type="POIUtility.invalidParameter", message="When specified, HeaderRowStart cannot be used without HasHeaderRow");
			}
			if( Arguments.HeaderRowStart LT 1 ){
				throw(type="POIUtility.rangeException", message="When specified, HeaderStart row must be greater than 0");
			}
			if(  StructKeyExists(ARGUMENTS,"DataRowStart") AND Arguments.HeaderRowStart GT Arguments.DataRowStart ){
				throw(type="POIUtility.rangeException", message="Header row cannot be after data row");
			}				
		}else{
			ARGUMENTS.HeaderRowStart = 1;
		}
		//we may not have a header where we are reading in columnnames, but we can grab the data anywere
		if( StructKeyExists(ARGUMENTS,"DataRowStart")  ){
			if( Arguments.DataRowStart LT 1 ){
				throw(type="POIUtility.rangeException", message="DataRowStart row must be greater than 0");
			}

		}

		if( StructKeyExists(ARGUMENTS,"ColumnStart")  ){
			if( Arguments.ColumnStart LT 1 ){
				throw(type="POIUtility.rangeException", message="ColumnStart column must be greater than 0");
			}

		}

		//We do the conversion to 0 based here
		if( ARGUMENTS.HasHeaderRow ){
			if(  StructKeyExists( ARGUMENTS,"HeaderRowStart" ) ){
				ARGUMENTS.HeaderRowStart -= 1;
			}else{
				ARGUMENTS.HeaderRowStart = 0;
			}
		}

		if( StructKeyExists( ARGUMENTS,"DataRowStart" ) ){
				ARGUMENTS.DataRowStart -= 1;
		}else{
			if( ARGUMENTS.HasHeaderRow ){
				ARGUMENTS.DataRowStart = ARGUMENTS.HeaderRowStart + 1;
			}else{
				ARGUMENTS.DataRowStart = 0;
			}
		}

		
		if( StructKeyExists( ARGUMENTS,"ColumnStart" ) ){
				ARGUMENTS.ColumnStart -= 1;
		}else{
			ARGUMENTS.ColumnStart = 0;
		}
		
		LOCAL.WBObjects = GetPOIWorkBookObjects( FilePath = ARGUMENTS.FilePath, type="read" );
		
		// Check to see if we are returning an array of sheets OR just
		// a given sheet.
		
		if (ARGUMENTS.SheetIndex GTE 0){

			// Read the sheet data for a single sheet.
			LOCAL.Sheets = ReadExcelSheet(
				Workbook          = ARGUMENTS.WorkBook,
				HasHeaderRow      = ARGUMENTS.HasHeaderRow,
				SheetIndex        = ARGUMENTS.SheetIndex,
				RowsToRead        = ARGUMENTS.RowsToRead,
				ColumnsToRead       = ARGUMENTS.ColumnsToRead,
				ColumnStart       = ARGUMENTS.ColumnStart,
				HeaderRowStart    = ARGUMENTS.HederRowStart,
				DataRowStart      = ARGUMENTS.DataRowStart	
				);
			
		} else {

			// No specific sheet was requested. We are going to return an array
			// of sheets within the Excel document.

			// Create an array to return.
			LOCAL.Sheets = ArrayNew( 1 );

			// Loop over the sheets in the document.
			for (
				LOCAL.SheetIndex = 0 ;
				LOCAL.SheetIndex LT LOCAL.WBObjects.WorkBook.GetNumberOfSheets() ;
				LOCAL.SheetIndex = (LOCAL.SheetIndex + 1)
				){

				// Add the sheet information.
				
				ArrayAppend(
					LOCAL.Sheets,
					ReadExcelSheet(
						Workbook          = LOCAL.WBObjects.WorkBook,
						HasHeaderRow      = ARGUMENTS.HasHeaderRow,
						SheetIndex        = LOCAL.SheetIndex,
						RowsToRead        = ARGUMENTS.RowsToRead,
						ColumnsToRead       = ARGUMENTS.ColumnsToRead,
						ColumnStart       = ARGUMENTS.ColumnStart,
						HeaderRowStart    = ARGUMENTS.HeaderRowStart,
						DataRowStart      = ARGUMENTS.DataRowStart	
						)
					);					
			}

		}
			
		// Return the array of sheets.
		return( LOCAL.Sheets );

	}

	/**
	* @hint Takes an Excel workbook and reads the given sheet (by index) into a structure
	* @WorkBook This is a workbook object created by the POI API.
	* @SheetIndex This is the index of the sheet within the passed in workbook. This is a ZERO-based index (coming from a Java object).
	* @HasHeaderRow Flags the Excel files has using the first data row a header column. If so, this column will be excluded from the resultant query.
	* @SheetIndex
	* @RowsToRead 
	* @ColumnsToRead 
	* @ColumnStart 
	* @HeaderRowStart
	* @DataRowStart 
	* @output false
	*/
	private struct function ReadExcelSheet( required any WorkBook, boolean HasHeaderRow = false, 
		numeric SheetIndex,
		numeric RowsToRead,
		numeric ColumnsToRead,
		numeric ColumnStart,
	    numeric HeaderRowStart,
	    numeric DataRowStart
		 ){

		// Set up the default return structure.
		LOCAL.SheetData = {};

		// This is the index of the sheet within the workbook.
		LOCAL.SheetData.Index = ARGUMENTS.SheetIndex;

		// This is the name of the sheet tab.
		LOCAL.SheetData.Name = ARGUMENTS.WorkBook.GetSheetName(
				JavaCast( "int", ARGUMENTS.SheetIndex )
				);

		// This is the query created from the sheet.
		LOCAL.SheetData.Query = QueryNew( "" );

		// This is a flag for the header row.
		LOCAL.SheetData.HasHeaderRow = ARGUMENTS.HasHeaderRow;

		// An array of header columns names.
		LOCAL.SheetData.ColumnNames = [];

		// This keeps track of the min number of data columns.
		LOCAL.SheetData.MinColumnCount = 0;

		// This keeps track of the max number of data columns.
		LOCAL.SheetData.MaxColumnCount = 0;

		// Get the sheet object at this index of the
		// workbook. This is based on the passed in data.
		LOCAL.Sheet = ARGUMENTS.WorkBook.GetSheetAt(
			JavaCast( "int", ARGUMENTS.SheetIndex )
		);

		//This can give a false value if a given row used to have data.  
		LOCAL.SheetData.MaxRowCount = LOCAL.Sheet.getLastRowNum();
		
		LOCAL.startRow = ( ARGUMENTS.HasHeaderRow  ? ARGUMENTS.HeaderRowStart:ARGUMENTS.DataRowStart );
		LOCAL.lastRow = ( ARGUMENTS.RowsToRead GT 0 AND (ARGUMENTS.DataRowStart + ARGUMENTS.RowsToRead ) LT LOCAL.Sheet.GetLastRowNum() ? ARGUMENTS.RowsToRead + ARGUMENTS.DataRowStart : LOCAL.Sheet.GetLastRowNum() );
		
		for (
			LOCAL.RowIndex = LOCAL.startRow;
			LOCAL.RowIndex LTE LOCAL.lastRow;
			LOCAL.RowIndex = (LOCAL.RowIndex + 1)
			){
				
			// Get a reference to the current row.
			LOCAL.Row = LOCAL.Sheet.GetRow(
				JavaCast( "int", LOCAL.RowIndex )
				);

			// Check to see if we are at an undefined row. If we are, then
			// our ROW variable has been destroyed.
			if (StructKeyExists( LOCAL, "Row" )){

				// Get the number of the last cell in the row. Since we
				// are in a defined row, we know that we must have at
				// least one row cell defined (and therefore, we must have
				// a defined cell number).
				LOCAL.ColumnCount = LOCAL.Row.GetLastCellNum();

				// Update the running min column count.
				LOCAL.SheetData.MinColumnCount = Min(
					LOCAL.SheetData.MinColumnCount,
					LOCAL.ColumnCount
					);

				// Update the running max column count.
				LOCAL.SheetData.MaxColumnCount = Max(
					LOCAL.SheetData.MaxColumnCount,
					LOCAL.ColumnCount
					);

			}

		}
		
		LOCAL.startCol = ARGUMENTS.ColumnStart;
		LOCAL.LastCol = ( ARGUMENTS.ColumnsToRead GT 0 AND (LOCAL.startCol + ARGUMENTS.ColumnsToRead ) LT LOCAL.SheetData.MaxColumnCount ? LOCAL.startCol + ARGUMENTS.ColumnsToRead : LOCAL.SheetData.MaxColumnCount );
	
		//Do a hard check on our range we intend to load
		
		LOCAL.lastRow = findLastRowWithData( Sheet=LOCAL.Sheet, 
		RowStart=LOCAL.startRow, RowEnd=LOCAL.lastRow,
		ColStart=LOCAL.startCol, ColEnd = LOCAL.LastCol);

		if ( LOCAL.lastRow eq -1 ){
			//Sheet is just empty
			return LOCAL.SheetData;
		}


		// ASSERT: At this pointer, we have a properly defined
		// query that will be able to handle any standard row
		// data that we encouter.


		// Loop over the rows in the Excel sheet. This time, we
		// already have a query built, so we just want to start
		// capturing the cell data.
		//Convert DataRowStart to 0 based
		for (
			LOCAL.RowIndex =  LOCAL.startRow;
			LOCAL.RowIndex LTE LOCAL.lastRow;
			LOCAL.RowIndex = ( LOCAL.RowIndex + 1 )
			){

			// Get a reference to the target row.
			LOCAL.Row = LOCAL.Sheet.GetRow(
				JavaCast( "int", LOCAL.RowIndex )
				);

			//Assumption, HEADER Row will ALWAYS come first if  StructKeyExists
			if( Arguments.HasHeaderRow AND LOCAL.RowIndex EQ Arguments.HeaderRowStart ){
				for (
						LOCAL.ColumnIndex = LOCAL.startCol;
						LOCAL.ColumnIndex LT LOCAL.LastCol ;
						LOCAL.ColumnIndex = (LOCAL.ColumnIndex + 1)
						){
						
						LOCAL.ColumnName = LOCAL.Row.GetCell( JavaCast( "int", LOCAL.ColumnIndex ) ).GetStringCellValue();
						ArrayAppend(LOCAL.SheetData.COLUMNNAMES, LOCAL.ColumnName );
						QueryAddColumn(
							LOCAL.SheetData.Query,
							LOCAL.ColumnName,
							"CF_SQL_VARCHAR",
							ArrayNew(1)
							);
					}
					continue;
			}else if ( ARGUMENTS.HasHeaderRow AND LOCAL.RowIndex GT ARGUMENTS.HeaderRowStart AND LOCAL.RowIndex LT ARGUMENTS.DataRowStart ){
				//Skip the between
				continue;
			}else if( ! ARGUMENTS.HasHeaderRow AND LOCAL.RowIndex EQ Arguments.DataRowStart ){
				//no header row so create the dummy columns and add another empty row
				for (
						LOCAL.ColumnIndex = LOCAL.startCol;
						LOCAL.ColumnIndex LT LOCAL.lastCol;
						LOCAL.ColumnIndex = (LOCAL.ColumnIndex + 1)
						){
						
						LOCAL.ColumnName = "column#LOCAL.ColumnIndex + LOCAL.startCol#";
						ArrayAppend(LOCAL.SheetData.COLUMNNAMES, LOCAL.ColumnName );
						QueryAddColumn(
							LOCAL.SheetData.Query,
							LOCAL.ColumnName,
							"CF_SQL_VARCHAR",
							ArrayNew(1)
							);
					}
			}

			QueryAddRow( LOCAL.SheetData.Query );

			// Check to see if we have a row. If we requested an
			// undefined row, then the NULL value will have
			// destroyed our Row variable.
			if ( StructKeyExists( LOCAL, "Row" ) ){

				// Get the number of the last cell in the row. Since we
				// are in a defined row, we know that we must have at
				// least one row cell defined (and therefore, we must have
				// a defined cell number).
				LOCAL.ColumnCount = LOCAL.Row.GetLastCellNum();

				// Now that we have an empty query, we are going to loop over
				// the cells COUNT for this data row and for each cell, we are
				// going to create a query column of type VARCHAR. I understand
				// that cells are going to have different data types, but I am
				// chosing to store everything as a string to make it easier.

				for (
					LOCAL.ColumnIndex =  LOCAL.startCol;
					LOCAL.ColumnIndex LT LOCAL.lastCol;
					LOCAL.ColumnIndex =  (LOCAL.ColumnIndex + 1)
					){
					
					// Check to see if we might be dealing with a header row.
					// This will be true if we are in the first row AND if
					// the user had flagged the header row usage.
					if ( ARGUMENTS.HasHeaderRow AND LOCAL.RowIndex EQ  ( ARGUMENTS.HeaderRowStart  ) ){

						
						// Try to get a header column name (it might throw
						// an error). We want to take that cell value and
						// add it to the array of header values that we will
						// return with the sheet data.
						try {

							// Add the cell value to the column names.
												
							ArrayAppend(
								LOCAL.SheetData.ColumnNames,
								LOCAL.Row.GetCell( 
								JavaCast( "int", LOCAL.ColumnIndex ) ).GetStringCellValue()
								);
							
						} catch (any ErrorHeader){

							// There was an error grabbing the text of the
							// header column type. Just add an empty string
							// to make up for it.
							ArrayAppend(
								LOCAL.SheetData.ColumnNames,
								""
								);
						}

					// We are either not using a Header row or we are no
					// longer dealing with the first row. In either case,
					// this data is standard cell data.
					} else {

						// When getting the value of a cell, it is important to know
						// what type of cell value we are dealing with. If you try
						// to grab the wrong value type, an error might be thrown.
						// For that reason, we must check to see what type of cell
						// we are working with. These are the cell types and they
						// are constants of the cell object itself:
				 		//
						// 0 - CELL_TYPE_NUMERIC
						// 1 - CELL_TYPE_STRING
						// 2 - CELL_TYPE_FORMULA
						// 3 - CELL_TYPE_BLANK
						// 4 - CELL_TYPE_BOOLEAN
						// 5 - CELL_TYPE_ERROR

						// Get the cell from the row object.
						LOCAL.Cell = LOCAL.Row.GetCell(
							JavaCast( "int", LOCAL.ColumnIndex )
							);

						// Check to see if we are dealing with a valid cell value.
						// If this was an undefined cell, the GetCell() will
						// have returned NULL which will have killed our Cell
						// variable.
						if (StructKeyExists( LOCAL, "Cell" )){

							// ASSERT: We are definitely dealing with a valid
							// cell which has some sort of defined value.

							// Get the type of data in this cell.

							//4.2 getCellType is deprecated
							LOCAL.CellType = LOCAL.Cell.GetCellType();

							// Get the value of the cell based on the data type. The thing
							// to worry about here is cell forumlas and cell dates. Formulas
							// can be strange and dates are stored as numeric types. For
							// this demo, I am not going to worry about that at all. I will
							// just grab dates as floats and formulas I will try to grab as
							// numeric values.
					
							if (LOCAL.CellType EQ LOCAL.CellType.NUMERIC) {

								// Get numeric cell data. This could be a standard number,
								// could also be a date value. I am going to leave it up to
								// the calling program to decide.
								LOCAL.CellValue = LOCAL.Cell.GetNumericCellValue();

							} else if (LOCAL.CellType EQ LOCAL.CellType.STRING){

								LOCAL.CellValue = LOCAL.Cell.GetStringCellValue();

							} else if (LOCAL.CellType EQ LOCAL.CellType.FORMULA){

								// Since most forumlas deal with numbers, I am going to try
								// to grab the value as a number. If that throws an error, I
								// will just grab it as a string value.
								try {

									LOCAL.CellValue = LOCAL.Cell.GetNumericCellValue();

								} catch (any Error1){

									// The numeric grab failed. Try to get the value as a
									// string. If this fails, just force the empty string.
									try {

										LOCAL.CellValue = LOCAL.Cell.GetStringCellValue();

									} catch (any Error2){

										// Force empty string.
										LOCAL.CellValue = "";

					 				}
								}

							} else if (LOCAL.CellType EQ LOCAL.CellType.BLANK){

								LOCAL.CellValue = "";

							} else if (LOCAL.CellType EQ LOCAL.CellType.BOOLEAN){

								LOCAL.CellValue = LOCAL.Cell.GetBooleanCellValue();

							} else {

								// If all else fails, get empty string.
								LOCAL.CellValue = "";

							}

							// ASSERT: At this point, we either got the cell value out of the
							// Excel data cell or we have thrown an error or didn't get a
							// matching type and just have the empty string by default.
							// No matter what, the object LOCAL.CellValue is defined and
							// has some sort of SIMPLE ColdFusion value in it.


							// Now that we have a value, store it as a string in the ColdFusion
							// query object. Remember again that my query names are ONE based
							// for ColdFusion standards. That is why I am adding 1 to the
							// cell index.

							//We need to subtract the local.startCol from the index 
							LOCAL.adjColumnIndex = LOCAL.ColumnIndex + 1 - Local.startCol;
							try{
								LOCAL.SheetData.Query[ LOCAL.SheetData.COLUMNNAMES[ LOCAL.adjColumnIndex ] ][ LOCAL.SheetData.Query.RecordCount  ] = JavaCast( "string", LOCAL.CellValue );
							}catch(any e){
								//writeDump(LOCAL);
								//WriteDump(e);
							}
						}

					}

				} 
				
			}
			
		}

		// Return the sheet object that contains all the Excel data.
		return( LOCAL.SheetData );
	}

	/**
	* @hint Takes an array of 'Sheet' structure objects and writes each of them to a tab in the Excel file.
	* @FilePath This is the expanded path of the Excel file.
	* @Sheets This is an array of the data that is needed for each sheet of the excel OR it is a single Sheet object. Each 'Sheet' will be a structure containing the Query, ColumnList, ColumnNames, and SheetName.
	* @Delimiters The list of delimiters used for the column list and column name arguments.
	* @HeaderCSS Defines the limited CSS available for the header row (if a header row is used).
	* @RowCSS Defines the limited CSS available for the non-header rows.
	* @AltRowCSS Defines the limited CSS available for the alternate non-header rows. This style overwrites parts of the RowCSS.
	* @output false
	*/
	public void function WriteExcel(required string FilePath, required any Sheets,
		string Delimiters=",", string HeaderCSS="", string RowCSS="", string AltRowCSS="" ){

		//TODO Javaloader and create workbook
		// Create Excel workbook.
		LOCAL.WBObjects = GetPOIWorkBookObjects( FilePath=ARGUMENTS.FilePath, type="write", createCSSRule=true );


		// Check to see if we are dealing with an array of sheets or if we were
		// passed in a single sheet.
		if (IsArray( ARGUMENTS.Sheets )){

			// This is an array of sheets. We are going to write each one of them
			// as a tab to the Excel file. Loop over the sheet array to create each
			// sheet for the already created workbook.
			for (
				LOCAL.SheetIndex = 1 ;
				LOCAL.SheetIndex LTE ArrayLen( ARGUMENTS.Sheets ) ;
				LOCAL.SheetIndex = (LOCAL.SheetIndex + 1)
				){


				// Create sheet for the given query information..
				WriteExcelSheet(
					WorkBook = LOCAL.WBObjects.WorkBook,
					CSSRule =  LOCAL.WBObjects.CSSRule,
					Query = ARGUMENTS.Sheets[ LOCAL.SheetIndex ].Query,
					ColumnList = ARGUMENTS.Sheets[ LOCAL.SheetIndex ].ColumnList,
					ColumnNames = ARGUMENTS.Sheets[ LOCAL.SheetIndex ].ColumnNames,
					SheetName = ARGUMENTS.Sheets[ LOCAL.SheetIndex ].SheetName,
					Delimiters = ARGUMENTS.Delimiters,
					HeaderCSS = ARGUMENTS.HeaderCSS,
					RowCSS = ARGUMENTS.RowCSS,
					AltRowCSS = ARGUMENTS.AltRowCSS
					);

			}

		} else {

			// We were passed in a single sheet object. Write this sheet as the
			// first and only sheet in the already created workbook.
			WriteExcelSheet(
				WorkBook = LOCAL.WBObjects.WorkBook,
				CSSRule =  LOCAL.WBObjects.CSSRule,
				Query = ARGUMENTS.Sheets.Query,
				ColumnList = ARGUMENTS.Sheets.ColumnList,
				ColumnNames = ARGUMENTS.Sheets.ColumnNames,
				SheetName = ARGUMENTS.Sheets.SheetName,
				Delimiters = ARGUMENTS.Delimiters,
				HeaderCSS = ARGUMENTS.HeaderCSS,
				RowCSS = ARGUMENTS.RowCSS,
				AltRowCSS = ARGUMENTS.AltRowCSS
				);

		}


		// ASSERT: At this point, either we were passed a single Sheet object
		// or we were passed an array of sheets. Either way, we now have all
		// of sheets written to the WorkBook object.


		// Create a file based on the path that was passed in. We will stream
		// the work data to the file via a file output stream.
		LOCAL.FileOutputStream = CreateObject(
			"java",
			"java.io.FileOutputStream"
			).Init(

				JavaCast(
					"string",
					ARGUMENTS.FilePath
					)

				);

		// Write the workout data to the file stream.
		LOCAL.WBObjects.WorkBook.Write(
			LOCAL.FileOutputStream
			);

		// Close the file output stream. This will release any locks on
		// the file and finalize the process.
		LOCAL.FileOutputStream.Close();

		// Return out.
		return;


	}

	/**
	* @hint Writes the given 'Sheet' structure to the given workbook
	* @WorkBook This is the Excel workbook that will create the sheets
	* @CSSRule  our very own CSSRule magic
	* @Query This is the query from which we will get the data
	* @ColumnList This is list of columns provided in custom-ordered
	* @ColumnNames This the the list of optional header-row column names. If this is not provided, no header row is used."
	* @SheetName This is the optional name that appears in this sheet's tab
	* @Delimiters The list of delimiters used for the column list and column name arguments.
	* @HeaderCSS Defines the limited CSS available for the header row (if a header row is used)
	* @RowCSS Defines the limited CSS available for the non-header rows
	* @AltRowCSS Defines the limited CSS available for the alternate non-header rows. This style overwrites parts of the RowCSS
	* output false
	*/
	public void function WriteExcelSheet(required any WorkBook, required any CSSRule, required any Query,
		                                string ColumnList=ARGUMENTS.Query.ColumnList,
		                                string ColumnNames="",
		                                string SheetName="Sheet #(ARGUMENTS.WorkBook.GetNumberOfSheets() + 1)#"
		                                string Delimiters=",", 
		                                string HeaderCSS="", string RowCSS="", string AltRowCSS=""){

			// Set up data type map so that we can map each column name to
			// the type of data contained.
			LOCAL.DataMap = {};

			// Get the meta data of the query to help us create the data mappings.
			LOCAL.MetaData = GetMetaData( ARGUMENTS.Query );

			// Loop over meta data values to set up the data mapping.
			for (
				LOCAL.MetaIndex = 1 ;
				LOCAL.MetaIndex LTE ArrayLen( LOCAL.MetaData ) ;
				LOCAL.MetaIndex = (LOCAL.MetaIndex + 1)
				){

				// Map the column name to the data type.
				LOCAL.DataMap[ LOCAL.MetaData[ LOCAL.MetaIndex ].Name ] = LOCAL.MetaData[ LOCAL.MetaIndex ].TypeName;
			}

			//default cell style
			LOCAL.Classes[ "@cell" ] = ARGUMENTS.CSSRule.AddCSS(
			StructNew(),
			ARGUMENTS.RowCSS
			) ;

			LOCAL.HeaderStyle = ARGUMENTS.WorkBook.CreateCellStyle();
			LOCAL.HeaderCSS = ARGUMENTS.CSSRule.AddCSS(LOCAL.Classes,ARGUMENTS.HeaderCSS );
			LOCAL.HeaderStyle = ARGUMENTS.CSSRule.ApplyToCellStyle(
				LOCAL.HeaderCSS,
				ARGUMENTS.Workbook,
				LOCAL.HeaderStyle
				);
 
 			LOCAL.RowStyle = ARGUMENTS.WorkBook.CreateCellStyle();
			LOCAL.RowCSS = ARGUMENTS.CSSRule.AddCSS(LOCAL.Classes,ARGUMENTS.RowCSS );
			LOCAL.RowStyle = ARGUMENTS.CSSRule.ApplyToCellStyle(
				LOCAL.RowCSS,
				ARGUMENTS.Workbook,
				LOCAL.RowStyle
				);

			LOCAL.AltRowStyle = ARGUMENTS.WorkBook.CreateCellStyle();
			LOCAL.AltRowCSS = ARGUMENTS.CSSRule.AddCSS(LOCAL.Classes,ARGUMENTS.AltRowCSS );

	
			// Now, loop over alt row css and check for values. If there are not
			// values (no length), then overwrite the alt row with the standard
			// row. This is a round-about way of letting the alt row override
			// the standard row.
			for (LOCAL.Key in LOCAL.AltRowCSS){

				// Check for value.
				if ( LOCAL.Key neq "@cell" AND NOT Len( LOCAL.AltRowCSS[ LOCAL.Key ] )){

					// Since we don't have an alt row style, copy over the standard
					// row style's value for this key.
					LOCAL.AltRowCSS[ LOCAL.Key ] = LOCAL.RowCSS[ LOCAL.Key ];

				}

			}

			LOCAL.AltRowStyle = ARGUMENTS.CSSRule.ApplyToCellStyle(
				LOCAL.AltRowCSS,
				ARGUMENTS.Workbook,
				LOCAL.AltRowStyle
				);


			// Create the sheet in the workbook.
			LOCAL.Sheet = ARGUMENTS.WorkBook.CreateSheet(
				JavaCast(
					"string",
					ARGUMENTS.SheetName
					)
				);

			// Set the sheet's default column width.
			LOCAL.Sheet.SetDefaultColumnWidth(
				JavaCast( "int", 23 )
				);


			// Set a default row offset so that we can keep add the header
			// column without worrying about it later.
			LOCAL.RowOffset = -1;

			// Check to see if we have any column names. If we do, then we
			// are going to create a header row with these names in order
			// based on the passed in delimiter.
			if (Len( ARGUMENTS.ColumnNames )){

				// Convert the column names to an array for easier
				// indexing and faster access.
				LOCAL.ColumnNames = ListToArray(
					ARGUMENTS.ColumnNames,
					ARGUMENTS.Delimiters
					);

				// Create a header row.
				LOCAL.Row = LOCAL.Sheet.CreateRow(
					JavaCast( "int", 0 )
					);

				// Set the row height.
				/*
				LOCAL.Row.SetHeightInPoints(
					JavaCast( "float", 14 )
					);
				*/


				// Loop over the column names.
				for (
					LOCAL.ColumnIndex = 1 ;
					LOCAL.ColumnIndex LTE ArrayLen( LOCAL.ColumnNames ) ;
					LOCAL.ColumnIndex = (LOCAL.ColumnIndex + 1)
					){

					// Create a cell for this column header.
					LOCAL.Cell = LOCAL.Row.CreateCell(
						JavaCast( "int", (LOCAL.ColumnIndex - 1) )
						);

					// Set the cell value.
					LOCAL.Cell.SetCellValue(
						JavaCast(
							"string",
							LOCAL.ColumnNames[ LOCAL.ColumnIndex ]
							)
						);

					// Set the header cell style.
					LOCAL.Cell.SetCellStyle(
						LOCAL.HeaderStyle
						);

				}

				// Set the row offset to zero since this will take care of
				// the zero-based index for the rest of the query records.
				LOCAL.RowOffset = 0;

			}

			// Convert the list of columns to the an array for easier
			// indexing and faster access.
			LOCAL.Columns = ListToArray(
				ARGUMENTS.ColumnList,
				ARGUMENTS.Delimiters
				);

			// Loop over the query records to add each one to the
			// current sheet.
			for (
				LOCAL.RowIndex = 1 ;
				LOCAL.RowIndex LTE ARGUMENTS.Query.RecordCount ;
				LOCAL.RowIndex = (LOCAL.RowIndex + 1)
				){

				// Create a row for this query record.
				LOCAL.Row = LOCAL.Sheet.CreateRow(
					JavaCast(
						"int",
						(LOCAL.RowIndex + LOCAL.RowOffset)
						)
					);

				/*
				// Set the row height.
				LOCAL.Row.SetHeightInPoints(
					JavaCast( "float", 14 )
					);
				*/


				// Loop over the columns to create the individual data cells
				// and set the values.
				for (
					LOCAL.ColumnIndex = 1 ;
					LOCAL.ColumnIndex LTE ArrayLen( LOCAL.Columns ) ;
					LOCAL.ColumnIndex = (LOCAL.ColumnIndex + 1)
					){

					// Create a cell for this query cell.
					LOCAL.Cell = LOCAL.Row.CreateCell(
						JavaCast( "int", (LOCAL.ColumnIndex - 1) )
						);

					// Get the generic cell value (short hand).
					LOCAL.CellValue = ARGUMENTS.Query[
						LOCAL.Columns[ LOCAL.ColumnIndex ]
						][ LOCAL.RowIndex ];

					// Check to see how we want to set the value. Meaning, what
					// kind of data mapping do we want to apply? Get the data
					// mapping value.
					LOCAL.DataMapValue = LOCAL.DataMap[ LOCAL.Columns[ LOCAL.ColumnIndex ] ];

					// Check to see what value type we are working with. I am
					// not sure what the set of values are, so trying to keep
					// it general.
					if (REFindNoCase( "int", LOCAL.DataMapValue )){

						LOCAL.DataMapCast = "int";

					} else if (REFindNoCase( "long", LOCAL.DataMapValue )){

						LOCAL.DataMapCast = "long";

					} else if (REFindNoCase( "double|decimal|numeric", LOCAL.DataMapValue )){

						LOCAL.DataMapCast = "double";

					} else if (REFindNoCase( "float|real|date|time", LOCAL.DataMapValue )){

						LOCAL.DataMapCast = "float";

					} else if (REFindNoCase( "bit", LOCAL.DataMapValue )){

						LOCAL.DataMapCast = "boolean";

					} else if (REFindNoCase( "char|text|memo", LOCAL.DataMapValue )){

						LOCAL.DataMapCast = "string";

					} else if (IsNumeric( LOCAL.CellValue )){

						LOCAL.DataMapCast = "float";

					} else {

						LOCAL.DataMapCast = "string";

					}

					// Set the cell value using the data map casting that we
					// just determined and the value that we previously grabbed
					// (for short hand).
					//
					// NOTE: Only set the cell value if we have a length. This
					// will stop us from improperly attempting to cast NULL values.
					if (Len( LOCAL.CellValue )){

						LOCAL.Cell.SetCellValue(
							JavaCast(
								LOCAL.DataMapCast,
								LOCAL.CellValue
								)
							);
					}

					// Get a pointer to the proper cell style. Check to see if we
					// are in an alternate row.
					if (LOCAL.RowIndex MOD 2){

						// Set standard row style.
						LOCAL.Cell.SetCellStyle(
							LOCAL.RowStyle
							);

					} else {

						// Set alternate row style.
						LOCAL.Cell.SetCellStyle(
							LOCAL.AltRowStyle
							);

					}

				}

			}


			// Return out.
			return;


	}
	/**
	* @hint Write the given query to an Excel file.
	* @output false
	* @FilePath expanded path of the excel file
	* @query This is the query from which we will get the data for the Excel file
	* @ColumnList This is list of columns provided in custom-order
	* @ColumnNames This the the list of optional header-row column names. If this is not provided, no header row is used
	* @SheetName This is the optional name that appears in the first (and only) workbook tab.
	* @Delimiters The list of delimiters used for the column list and column name arguments.
	* @HeaderCSS Defines the limited CSS available for the header row (if a header row is used)
	* @RowCSS Defines the limited CSS available for the non-header rows.
	* @AltRowCSS Defines the limited CSS available for the alternate non-header rows. This style overwrites parts of the RowCSS
	*/
	public void function WriteSingleExcel( required string FilePath, required query Query
											, string ColumnList=ARGUMENTS.Query.ColumnList
											, string ColumnNames="", string Sheetname="Sheet 1"
											, string Delimiters=",", string HeaderCSS="", string RowCSS="", string AltRowCSS=""){
			
			

			// Get a new sheet object.
			LOCAL.Sheet = GetNewSheetStruct( ARGUMENTS );

			// Write this sheet to an Excel file.
			WriteExcel(
				FilePath = ARGUMENTS.FilePath,
				Sheets = LOCAL.Sheet,
				Delimiters = ARGUMENTS.Delimiters,
				HeaderCSS = ARGUMENTS.HeaderCSS,
				RowCSS = ARGUMENTS.RowCSS,
				AltRowCSS = ARGUMENTS.AltRowCSS
				);

			// Return out.
			return;


	}


	/**
	* @hint Returns a default structure of what this Component is expecting for a sheet definition when WRITING Excel files
	* @output false
	*/
	public struct function GetNewSheetStruct( ){

		LOCAL.result = {Query = "",ColumnList="",ColumnNames="",SheetName="" };
		
		for(LOCAL.prop in ARGUMENTS){
			if( StructKeyExists( LOCAL.result, LOCAL.prop ) ){
				LOCAL.result[prop] = ARGUMENTS[prop];
			}
		}
	
		return LOCAL.result;
	}

	public struct function Debug (){
		writeDump(ARGUMENTS);
		abort;
	}

	public numeric function findLastRowWithData(required any Sheet, 
		required numeric RowStart, required numeric RowEnd,
		required numeric ColStart, required numeric ColEnd ){

		LOCAL.hasData = false;
		/** Get last row is problematic in that it can see formerly used row as a last row.
		* the idea here is to take the value what POI thinks is a last row and then iterate
		* backwards to find the first non-null row and return that.
		*/
		for(local.rowIndex = ARGUMENTS.RowEnd; local.rowIndex GTE ARGUMENTS.RowStart; local.rowIndex -= 1 ){
			//sheet get row reference
			// Get a reference to the current row.
			LOCAL.Row = ARGUMENTS.Sheet.GetRow(
				JavaCast( "int", LOCAL.RowIndex )
				);

			LOCAL.hasCells = false;
			for( local.colIndex = ARGUMENTS.ColEnd; local.colIndex GTE ARGUMENTS.ColStart; local.colIndex -= 1){

				// Get the cell from the row object.
				LOCAL.Cell = LOCAL.Row.GetCell(
					JavaCast( "int", LOCAL.colIndex )
					);
				if (StructKeyExists( LOCAL, "Cell" )){
					
					//4.2 getCellType is deprecated
					LOCAL.CellType = LOCAL.Cell.GetCellType();
					if ( LOCAL.Cell.getCellType()  neq LOCAL.CellType.BLANK ){
						LOCAL.hasCells = true;
					}
					

				}
				
			}	
			if ( LOCAL.hasCells ){
				return local.rowIndex;
			}	
		}



		return -1;
	}

}