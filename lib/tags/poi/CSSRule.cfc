/**
*	CSSRule.cfc
* 	A revamp of Ben Nadel's CSSRule.cfc to support Apache POI 4.0x, updated to cfscript
*   Accessors are enabled to make it easier to write unit tests to grab the available values
* 
*   @author: Chris Wigginton
* 	@verison 4.0
* 
* 	@hint Handles CSS utility functions." 
*   @accessors true
* 	@output false
*/
component{

   property struct CSS;
   property struct CSSCache;
   property struct CSSValidation;
   property any    IndexedColorMap;
   property any    IndexedColors;
   property struct POIColors;
   property struct SortedPropertyKeys;
   property any    borderStyle;
   property struct borderStyles;
   property struct cssClasses;
   property any    fillPattern;
   property struct fillPatterns;
   property struct horizontalAlignments;
   property any    javaLoader;
   property any    regionUtil;
   property struct verticalAlignments;
   property any    workbook;

  	
	/**
	* @hint Returns an initialized component.
	* @javaLoader to configure this to use the latest and greates Apache POI and
	*			  and make it support both xlsx and xls
	* @workbook   to make the methods available from the workbook
	*/
	// TODO: Verify xls workbook has method getWorkbookType
	public any function Init(required any javaLoader, required any workbook ) output="false"{

		if( StructKeyExists(ARGUMENTS,"javaLoader") ){
			VARIABLES.javaLoader = ARGUMENTS.javaLoader;
		}

		if( StructKeyExists(ARGUMENTS,"workbook") ){
			VARIABLES.workbook = ARGUMENTS.workbook;
		}

		//TODO investigate the xssf clases to see if we can use the ss to support xls and xlsx
		if( arguments.workbook.getWorkBookType().getExtension() eq "xlsx" ){
			VARIABLES.isXLSX = true;
			VARIABLES.classes = {
				cellStyle            = "org.apache.poi.xssf.usermodel.XSSFCellStyle",
				borderStyle          = "org.apache.poi.ss.usermodel.BorderStyle",
				color                = "org.apache.poi.xssf.usermodel.XSSFColor",
				colorIndex           = "org.apache.poi.ss.usermodel.IndexedColors",
				fillPattern          = "org.apache.poi.ss.usermodel.FillPatternType",
				horizontalAlignment  = "org.apache.poi.ss.usermodel.HorizontalAlignment",
				verticalAlignment    = "org.apache.poi.ss.usermodel.VerticalAlignment",
				cellRangeAddress     = "org.apache.poi.ss.util.CellRangeAddress",
				RegionUtil           ="org.apache.poi.ss.util.RegionUtil"
			};
		}else {
			VARIABLES.isXLSX = false;
			VARIABLES.classes = {
				cellStyle = "org.apache.poi.hssf.usermodel.HSSFCellStyle",
				color     ="org.apache.poi.hssf.util.HSSFColor"
			};
		};
		
		
		// Keep for reference org.apache.poi.hssf.util.HSSFColor.HSSFColorPredefined

		//org.apache.poi.ss.usermodel.IndexedColors


		// Set up the default CSS properties for this rule. This will
		// be used to create other hash maps.
		VARIABLES.CSS = {
		"background-attachment" = "",
		"background-color" = "",
		"background-image" = "",
		"background-pattern" = "",
		"background-position" = "",
		"background-repeat" = "",
		"border-width" = "",
		"border-color" = "",
		"border-style" = "",
		"border-top-width" = "",
		"border-top-color" = "",
		"border-top-style" = "",
		"border-right-width" = "",
		"border-right-color" = "",
		"border-right-style" = "",
		"border-bottom-width" = "",
		"border-bottom-color" = "",
		"border-bottom-style" = "",
		"border-left-width" = "",
		"border-left-color" = "",
		"border-left-style" = "",
		"bottom" = "",
		"color" = "",
		"display" = "",
		"font-family" = "",
		"font-size" = "",
		"font-style" = "",
		"font-weight" = "",
		"height" = "",
		"left" = "",
		"list-style-image" = "",
		"list-style-position" = "",
		"list-style-type" = "",
		"margin" = "",
		"margin-top" = "",
		"margin-right" = "",
		"margin-bottom" = "",
		"margin-left" = "",
		"padding" = "",
		"padding-top" = "",
		"padding-right" = "",
		"padding-bottom" = "",
		"padding-left" = "",
		"position" = "",
		"right" = "",
		"text-align" = "",
		"text-decoration" = "",
		"top" = "",
		"vertical-align" = "",
		"white-space" = "",
		"width" = "",
		"z-index" = ""};

		/*
			Set up the validation rules for the CSS properties. Each
			property must fit in a certain format. These formats
			will be defined using regular expressions and will be
			used to match the entire value (no partial matching).

			TODO: update the regex for the generics and to match the current styles
		*/
		VARIABLES.CSSValidation = {
		"background-attachment" = "scroll|fixed",
		"background-color" = "\w+|##[0-9ABCDEF]{6}",
		"background-image" = "url\([^\)]+\)",
		"background-pattern" = ".*",
		"background-position" = "(top|right|bottom|left|\d+(\.\d+)?(px|%|em)) (top|right|bottom|left|\d+(\.\d+)?(px|%|em))",
		"background-repeat" = "(no-)?repeat(-x|-y)?",
		"border-top-width" = "\d+(\.\d+)?px",
		"border-top-color" = "\w+|##[0-9ABCDEF]{6}",
		"border-top-style" = "none|dotted|dashed|solid|double|groove",
		"border-right-width" = "\d+(\.\d+)?px",
		"border-right-color" = "\w+|##[0-9ABCDEF]{6}",
		"border-right-style" = "none|dotted|dashed|solid|double|groove",
		"border-bottom-width" = "\d+(\.\d+)?px",
		"border-bottom-color" = "\w+|##[0-9ABCDEF]{6}",
		"border-bottom-style" = "none|dotted|dashed|solid|double|groove",
		"border-left-width" = "\d+(\.\d+)?px",
		"border-left-color" = "\w+|##[0-9ABCDEF]{6}",
		"border-left-style" = "none|dotted|dashed|solid|double|groove",
		"bottom" = "-?\d+(\.\d+)?px",
		"color" = "\w+|##[0-9ABCDEF]{6}",
		"display" = "inline|block|block",
		"font-family" = "((\w+|""[^""]""+)(\s*,\s*)?)+",
		"font-size" = "\d+(\.\d+)?(px|pt|em|%)",
		"font-style" = "normal|italic",
		"font-weight" = "normal|lighter|bold|bolder|[1-9]00",
		"height" = "\d+(\.\d+)?(px|pt|em|%)",
		"left" = "-?\d+(\.\d+)?px",
		"list-style-image" = "none|url\([^\)]+\)",
		"list-style-position" = "inside|outside",
		"list-style-type" = "disc|circle|square|none",
		"margin-top" = "\d+(\.\d+)?(px|em)",
		"margin-right" = "\d+(\.\d+)?(px|em)",
		"margin-bottom" = "\d+(\.\d+)?(px|em)",
		"margin-left" = "\d+(\.\d+)?(px|em)",
		"padding-top" = "\d+(\.\d+)?(px|em)",
		"padding-right" = "\d+(\.\d+)?(px|em)",
		"padding-bottom" = "\d+(\.\d+)?(px|em)",
		"padding-left" = "\d+(\.\d+)?(px|em)",
		"position" = "static|relative|absolute|fixed",
		"right" = "-?\d+(\.\d+)?px",
		"text-align" = "left|right|center|justify",
		"text-decoration" = "none|underline|overline|line-through",
		"top" = "-?\d+(\.\d+)?px",
		"vertical-align" = "center|distributed|justify|bottom|top",
		"white-space" = "normal|pre|nowrap",
		"width" = "\d+(\.\d+)?(px|pt|em|%)|auto",
		"z-index" = "\d+"};


	// Here is an array of the alpha-sorted keys.
	VARIABLES.SortedPropertyKeys = StructKeyArray( VARIABLES.CSS );

	// Sort the keys alphabetically.
	ArraySort( VARIABLES.SortedPropertyKeys, "textnocase", "asc" );


		/*
			This is going to be a cached value of CSS strings. I am doing this
			because if someone has a "style" inside of a large loop, I don't want
			to be re-parsing that every single time.
		*/
		VARIABLES.CSSCache = {};

		// Create a struct of valid colors.
		VARIABLES.POIColors = {
			AQUA                  = true,
			AUTOMATIC             = true,
			BLACK                 = true,
			BLACK1                = true,
			BLUE                  = true,
			BLUE1                 = true,
			BLUE_GREY             = true,
			BRIGHT_GREEN          = true,
			BRIGHT_GREEN1         = true,
			BROWN                 = true,
			CORAL                 = true,
			CORNFLOWER_BLUE       = true,
			DARK_BLUE             = true,
			DARK_GREEN            = true,
			DARK_RED              = true,
			DARK_TEAL             = true,
			DARK_YELLOW           = true,
			GOLD                  = true,
			GREEN                 = true,
			GREY_25_PERCENT       = true,
			GREY_40_PERCENT       = true,
			GREY_50_PERCENT       = true,
			GREY_80_PERCENT       = true,
			INDIGO                = true,
			LAVENDER              = true,
			LEMON_CHIFFON         = true,
			LIGHT_BLUE            = true,
			LIGHT_CORNFLOWER_BLUE = true,
			LIGHT_GREEN           = true,
			LIGHT_ORANGE          = true,
			LIGHT_TURQUOISE       = true,
			LIGHT_TURQUOISE1      = true,
			LIGHT_YELLOW          = true,
			LIME                  = true,
			MAROON                = true,
			OLIVE_GREEN           = true,
			ORANGE                = true,
			ORCHID                = true,
			PALE_BLUE             = true,
			PINK                  = true,
			PINK1                 = true,
			PLUM                  = true,
			RED                   = true,
			RED1                  = true,
			ROSE                  = true,
			ROYAL_BLUE            = true,
			SEA_GREEN             = true,
			SKY_BLUE              = true,
			TAN                   = true,
			TEAL                  = true,
			TURQUOISE             = true,
			TURQUOISE1            = true,
			VIOLET                = true,
			WHITE                 = true,
			WHITE1                = true,
			YELLOW                = true,
			YELLOW1               = true
		};

		VARIABLES.borderStyles = {
			 DASHED              = true,
			 DASH_DOT	         = true,
			 DASH_DOT_DOT        = true,
			 DOTTED              = true,
			 DOUBLE	             = true,
			 HAIR	             = true,
			 MEDIUM	             = true,
			 MEDIUM_DASHED	     = true,
			 MEDIUM_DASH_DOT	 = true,
			 MEDIUM_DASH_DOT_DOT = true,
			 NONE	             = true,
			 SLANTED_DASH_DOT	 = true,
			 THICK	             = true,
			 THIN                = true
		};

		VARIABLES.fillPatterns = {
			ALT_BARS            = true,
			BIG_SPOTS           = true,
			BRICKS              = true,
			ALT_BARS	        = true,
			BIG_SPOTS	        = true,
			BRICKS	            = true,
			DIAMONDS	        = true,
			FINE_DOTS           = true,
			LEAST_DOTS	        = true,
			LESS_DOTS	        = true,
			NO_FILL	            = true,
			SOLID_FOREGROUND    = true,
			SPARSE_DOTS	        = true,
			SQUARES	            = true,
			THICK_BACKWARD_DIAG = true,
			THICK_FORWARD_DIAG	= true,
			THICK_HORZ_BANDS	= true,
			THICK_VERT_BANDS	= true,
			THIN_BACKWARD_DIAG	= true,
			THIN_FORWARD_DIAG	= true,
			THIN_HORZ_BANDS	    = true,
			THIN_VERT_BANDS     = true	
		};

		VARIABLES.horizontalAlignments = {
			CENTER           = true,
			CENTER_SELECTION = true,
			DISTRIBUTED      = true,
			FILL             = true,
			GENERAL          = true,
			JUSTIFY          = true,
			LEFT             = true,
			RIGHT            = true
		};

		
		VARIABLES.verticalAlignments = {
			BOTTOM           = true,
			CENTER           = true,
			DISTRIBUTED      = true,
			JUSTIFY          = true,
			TOP              = true
		};

	

		// other possibilities
		//org.apache.poi.hssf.util.HSSFColor.HSSFColorPredefined


		//Class Instances for constants
		VARIABLES.borderStyle = VARIABLES.javaLoader.create(  VARIABLES.classes.borderStyle );
		VARIABLES.fillPattern = VARIABLES.javaLoader.create(  VARIABLES.classes.fillPattern );
		VARIABLES.cellStyle = VARIABLES.javaLoader.create(  VARIABLES.classes.cellStyle );
		VARIABLES.IndexedColorMap = ARGUMENTS.workbook.getStylesSource().getIndexedColors();
		VARIABLES.IndexedColors = VARIABLES.javaLoader.create(  VARIABLES.classes.colorIndex );
		VARIABLES.horizontalAlignment = VARIABLES.javaLoader.create(  VARIABLES.classes.horizontalAlignment );
		VARIABLES.verticalAlignment = VARIABLES.javaLoader.create(  VARIABLES.classes.verticalAlignment );

		VARIABLES.regionUtil =VARIABLES.javaLoader.create(  VARIABLES.classes.regionUtil );

		return THIS;
	}

	/**
	* @hint Adds CSS properties to passed-in css hash map returns it.
	* @PropertyMap I am the CSS hash map being updated.
	* @CSS CSS properties for to be added to the given map (may have multiple properties separated by semi-colons).
	* 
	* @output false
	*/
	public struct function AddCSS( required struct PropertyMap, required string CSS){


		/*
			Check to see if this CSS string has already been cached. If not,
			then we want to cache it locally first, then add it to the struct.
		*/

		if(  NOT StructKeyExists( VARIABLES.CSSCache, ARGUMENTS.CSS ) ){

			// Create a local property map.
			LOCAL.CachedPropertyMap = {};

			LOCAL.CSSProps = ListToArray(ARGUMENTS.CSS,";");
			// Loop over the list of properties passed in.
			for( LOCAL.Property in LOCAL.CSSProps ){
				// Add this property to the css map.
				AddProperty(
					LOCAL.CachedPropertyMap,
					Trim( LOCAL.Property )
					);
			}

			// Cache this property map.
			VARIABLES.CSSCache[ ARGUMENTS.CSS ] = LOCAL.CachedPropertyMap;
		}


		/*
			ASSERT: At this point, we know that no matter what CSS string was
			passed-in, we now have a version of it parsed and stored in the cache.
		*/


		// Add the cached property map.
		StructAppend(
			ARGUMENTS.PropertyMap,
			VARIABLES.CSSCache[ ARGUMENTS.CSS ]
			);

		// Return the updated map.
		return ARGUMENTS.PropertyMap;
	}

	/**
	* @hint Parses the given property and adds it to the given CSS property map.
	* @PropertyMap I am the CSS hash map being updated
	* @Property The name-value pair property that will be added to the CSS rule.
	* 
	* @output false
	*/
	public boolean function AddProperty( required struct PropertyMap, required string Property ){


		/*
			The property should be in name=value pair format. Break up the
			property into the two parts. Also, make sure that we only have
			one property being set (as delimited by ";").
		*/
		LOCAL.Pair = ListToArray(
			Trim( ListFirst( ARGUMENTS.Property , ";" ) ),
			":"
			);

		/*
			Check to see if we have two parts. If we have
			anything but two parts, then this is not a valid
			name-value pair.
		*/
		if(ArrayLen( LOCAL.Pair ) EQ 2){

			// Trim both parts of the pair.
			LOCAL.Name = Trim( LOCAL.Pair[ 1 ] );
			LOCAL.Value = Trim( LOCAL.Pair[ 2 ] );

			/*
				When it comes to parsing the property, they might be
				using a simple one that we have. If not, we have to
				get a little more creative with the parsing.
			*/
			if( IsValidValue( LOCAL.Name, LOCAL.Value ) ){

				// This value has validated. Add it to the CSS properties.
				ARGUMENTS.PropertyMap[ LOCAL.Name ] = LOCAL.Value;

				// Return true for success.
				return true;

			}else{

				/*
					We were not given a simple value; we were given a value that
					we will have to parse out into the individual properties.
				*/
				LOCAL.propParent = ListGetAt(Local.Name,1,"-");
				switch (LOCAL.propParent){

					case "background":
						SetBackground( ARGUMENTS.PropertyMap, LOCAL.Value );
						break;

					case "border":
						SetBorder( ARGUMENTS.PropertyMap, LOCAL.Name, LOCAL.Value );
						break;

					case "font":
						SetFont( ARGUMENTS.PropertyMap, LOCAL.Value );
						break;

					case "list":
						SetListStyle( ARGUMENTS.PropertyMap, LOCAL.Value );
						break;

					case "margin":
						SetMargin( ARGUMENTS.PropertyMap, LOCAL.Value );
						break;

					case "padding":
						SetPadding( ARGUMENTS.PropertyMap, LOCAL.Value );
						break;
				}
			}
		}

		/*
			Return out. If we made it this far, then we
			didn't add a valid property.
		*/
		return false;
	}

	/**
	* @hint Applies the current CSS property map to the given HSSFCellStyle object
	* @PropertyMap I am the CSS hash map being updated.
	* @Workbook The workbook containing this cell style.
	* @CellStyle The HSSFCellStyle instance to which we are applying the CSS property rules.
	* 
	* @output false
	*/
	public any function ApplyToCellStyle(required struct PropertyMap, required any Workbook, required any CellStyle ){

		// Create a local copy of the full CSS definition.
		LOCAL.PropertyMap = StructCopy( VARIABLES.CSS );

		/*
			Now, append the passed in property map to this local one. That will give
			us a full CSS property map with only the relatvant values filled in.
		*/
		StructAppend( LOCAL.PropertyMap, ARGUMENTS.PropertyMap );

		// Get a new font object from the workbook.
		LOCAL.Font = ARGUMENTS.WorkBook.CreateFont();

		if( Len( LOCAL.PropertyMap[ "background-color" ] )
			AND StructKeyExists( VARIABLES.POIColors, LOCAL.PropertyMap[ "background-color" ] ) ){
		 	ARGUMENTS.CellStyle.SetFillForegroundColor( getXSSFColorByName( UCase( LOCAL.PropertyMap[ 'background-color' ] ) ) );

			 //let background-pattern do an override
			if (! Len( LOCAL.PropertyMap[ "background-pattern" ] ) ){
				ARGUMENTS.CellStyle.SetFillPattern( VARIABLES.fillPattern.SOLID_FOREGROUND );
			}
		}

		if( LOCAL.PropertyMap[ "background-color" ] EQ "transparent" ){
			// The user has requested no background color.
			ARGUMENTS.CellStyle.SetFillPattern( VARIABLES.fillPattern.NO_FILL );
		}else if( Len( LOCAL.PropertyMap[ "background-pattern" ] ) ){

			if( StructKeyExists( VARIABLES.fillPatterns, UCase( LOCAL.PropertyMap[ "background-pattern" ] ) ) ){
				ARGUMENTS.CellStyle.SetFillPattern( evaluate("VARIABLES.fillPattern.#UCase( LOCAL.PropertyMap[ 'background-pattern' ] )#") );
			}
		}


		//handle generic border-style
		if(  Len( LOCAL.PropertyMap[ "border-style" ] ) ){
			if( StructKeyExists( VARIABLES.borderStyles, Ucase( LOCAL.PropertyMap[ 'border-style' ] ) ) ){
					LOCAL.BorderStyle = evaluate("VARIABLES.BorderStyle.#LOCAL.PropertyMap[ 'border-style' ]#");
					ARGUMENTS.CellStyle.SetBorderTop( LOCAL.BorderStyle );
					ARGUMENTS.CellStyle.SetBorderBottom( LOCAL.BorderStyle );
					ARGUMENTS.CellStyle.SetBorderRight( LOCAL.BorderStyle );
					ARGUMENTS.CellStyle.SetBorderLeft( LOCAL.BorderStyle );
				}
		}
		//handle generic border color

		if(  Len( LOCAL.PropertyMap[ "border-color" ] ) ){

			if( StructKeyExists( VARIABLES.POIColors, Ucase( LOCAL.PropertyMap[ 'border-color' ] ) ) ){
				    LOCAL.BorderColor = getXSSFColorByName( UCase( PropertyMap[ 'border-color' ] ) );
				    
					ARGUMENTS.CellStyle.SetTopBorderColor( LOCAL.BorderColor );
					ARGUMENTS.CellStyle.SetRightBorderColor( LOCAL.BorderColor );
					ARGUMENTS.CellStyle.SetBottomBorderColor( LOCAL.BorderColor );
					ARGUMENTS.CellStyle.SetLeftBorderColor( LOCAL.BorderColor );
				}
		}

		// Loop over the four border directions.
		LOCAL.borderDirections = ["top","right","bottom","left"];
		for( LOCAL.BorderSide in LOCAL.borderDirections ){
			//set directional border styles
			if(  Len( LOCAL.PropertyMap[ "border-#LOCAL.BorderSide#-style" ] )
			  AND  StructKeyExists( VARIABLES.borderStyles, Ucase( LOCAL.PropertyMap[ 'border-#LOCAL.BorderSide#-style' ] ) ) ){
				LOCAL.BorderStyle = evaluate("VARIABLES.BorderStyle.#LOCAL.PropertyMap[ 'border-#LOCAL.BorderSide#-style' ]#");
				// Check to see which direction we are working width.
				switch ("#LOCAL.BorderSide#"){
					case "top":
						ARGUMENTS.CellStyle.SetBorderTop( LOCAL.BorderStyle );
						break;
					case "right":
						ARGUMENTS.CellStyle.SetBorderRight( LOCAL.BorderStyle );
						break;
					case "bottom":
						ARGUMENTS.CellStyle.SetBorderBottom( LOCAL.BorderStyle );
						break;
					case "left":
						ARGUMENTS.CellStyle.SetBorderLeft( LOCAL.BorderStyle );
						break;
				}
			}

			//set directional border colors
			if(  Len( LOCAL.PropertyMap[ "border-#LOCAL.BorderSide#-color" ] ) 
				AND StructKeyExists( VARIABLES.POIColors, Ucase( LOCAL.PropertyMap[ 'border-#LOCAL.BorderSide#-color' ] ) ) ){
				LOCAL.BorderColor = getXSSFColorByName( UCase( PropertyMap[ 'border-#LOCAL.BorderSide#-color' ] ) );
				switch ("#LOCAL.BorderSide#"){
					case "top":
						ARGUMENTS.CellStyle.SetTopBorderColor( LOCAL.BorderColor );
						break;
					case "right":
						ARGUMENTS.CellStyle.SetRightBorderColor( LOCAL.BorderColor );
						break;
					case "bottom":
						ARGUMENTS.CellStyle.SetBottomBorderColor( LOCAL.BorderColor );
						break;
					case "left":
						ARGUMENTS.CellStyle.SetLeftBorderColor( LOCAL.BorderColor );
						break;
				}
			}
		} // end border directions


		/*
			Check to see if we have an appropriate text color; Excel will not
			just use any color - it has to be one of their index colors.
		*/
		if(
			Len( LOCAL.PropertyMap[ "color" ] ) AND
			StructKeyExists( VARIABLES.POIColors, LOCAL.PropertyMap[ "color" ] )
			){

			LOCAL.Font.SetColor( getXSSFColorByName( UCase( LOCAL.PropertyMap[ "color" ] ) ) );
		}

		// Check for font family.
		if(  Len( LOCAL.PropertyMap[ "font-family" ] ) ){
			LOCAL.Font.SetFontName(
				JavaCast( "string", LOCAL.PropertyMap[ "font-family" ] )
				);
		}

		// Check for font style.
		switch ("#LOCAL.PropertyMap[ 'font-style' ]#"){
			case "italic":
				LOCAL.Font.SetItalic( JavaCast( "boolean", true ) );
				break;
		}

		// Check for font weight.
		switch ("#LOCAL.PropertyMap[ 'font-weight' ]#"){

			case "bold":
			case "600":
			case "700":
			case "800":
			case "900":
				LOCAL.Font.SetBoldWeight(
					LOCAL.Font.BOLDWEIGHT_BOLD
					);
				break;
			case "normal":
			case "100":
			case "200":
			case "300":
			case "400":
			case "500":
				LOCAL.Font.SetBoldWeight(
					LOCAL.Font.BOLDWEIGHT_NORMAL
					);
				break;
		}


		// Check for font size.
		if(  Val( LOCAL.PropertyMap[ "font-size" ] ) ){
			LOCAL.Font.SetFontHeightInPoints(
				JavaCast( "int", Val( LOCAL.PropertyMap[ "font-size" ] ) )
				);
		}


		// Check to see if we have any text alignment.
		if( StructKeyExists( VARIABLES.horizontalAlignments, ucase( LOCAL.PropertyMap[ 'text-align' ] )  ) ){
			ARGUMENTS.CellStyle.SetAlignment( Evaluate("VARIABLES.horizontalAlignment.#ucase( LOCAL.PropertyMap[ 'text-align' ] )#" ) );
		}
	

		// Check to see if we have any vertical alignment.
		if( StructKeyExists( VARIABLES.verticalAlignments, ucase( LOCAL.PropertyMap[ 'vertical-align' ] )  ) ){
			ARGUMENTS.CellStyle.SetVerticalAlignment( Evaluate("VARIABLES.verticalAlignment.#ucase( LOCAL.PropertyMap[ 'vertical-align' ] )#" ) );
		}
	
		/*
			Check for white space. If we have normal, which is the default, then
			let's turn on the text wrap. If we have anything else, then turn off
			the text wrap.
		*/

		switch("#LOCAL.PropertyMap[ 'white-space' ]#"){
			case "nowrap":
			case "pre":
				ARGUMENTS.CellStyle.SetWrapText( JavaCast( "boolean", false ) );
				break;
			default:
				// Default is "normal", which will turn it on.
				ARGUMENTS.CellStyle.SetWrapText( JavaCast( "boolean", true ) );
				break;
		}

		// Apply the font to the current style.
		ARGUMENTS.CellStyle.SetFont( LOCAL.Font );

		// Return the updated cell style object.
		return ARGUMENTS.CellStyle;
	}


	/**
	* @hint Parses the property value into individual tokens
	* @Value The value we want to parse into an array of tokens.
	*/
	public array function GetPropertyTokens( required string Value ){

		/*
			Get the tokens. These are the smallest meaningful
			pieces of any CSS property.
		*/
		return REMatch(
			(
				"(?i)" &
				"url\([^\)]+\)|" &
				"""[^""]+""|" &
				"##[0-9ABCDEF]{6}|" &
				"([\w\.\-%]+(\s*,\s*)?)+"
			),
			ARGUMENTS.Value
			);
	}

	/**
	* @hint Checks to see if the given value validated for a given property.
	* @Property The property we are checking for.
	* @Value The value we are checking for validity
	* 
	* @output false
	*/
	public boolean function IsValidValue( required string Property, required string Value ){

		/*
			Return whether it validates. If the property is not
			valid, we are returning false (same as an invalid value).
		*/
		return true;
		return (
			StructKeyExists( VARIABLES.CSS, ARGUMENTS.Property ) AND
			REFind( "(?i)^#VARIABLES.CSSValidation[ ARGUMENTS.Property ]#$", ARGUMENTS.Value )
			);
	}

	/**
	* @hint Takes a quad metric and returns a four-point array.
	* @Value The metric which may have between one and four values.
	* 
	* @output false
	*/
	public array function ParseQuadMetric( required string Value ){

		// Grab metric values.
		LOCAL.Values = REMatch( "\d+(\.\d+)?(px|em)", ARGUMENTS.Value );

		// Set up the return array.
		LOCAL.Return = [ "", "", "", "" ];

		// Check to see how many values we have.
		if(ArrayLen( LOCAL.Values ) EQ 1){

			// Copy to all positions.
			ArraySet( LOCAL.Return, 1, 4, LOCAL.Values[ 1 ] );

		}else if(ArrayLen( LOCAL.Values ) EQ 2){

			// Copy 2 and 2.
			LOCAL.Return[ 1 ] = LOCAL.Values[ 1 ];
			LOCAL.Return[ 2 ] = LOCAL.Values[ 2 ];
			LOCAL.Return[ 3 ] = LOCAL.Values[ 1 ];
			LOCAL.Return[ 4 ] = LOCAL.Values[ 2 ];

		}else if(ArrayLen( LOCAL.Values ) EQ 3){

			// Copy 3 and 1.
			LOCAL.Return[ 1 ] = LOCAL.Values[ 1 ];
			LOCAL.Return[ 2 ] = LOCAL.Values[ 2 ];
			LOCAL.Return[ 3 ] = LOCAL.Values[ 3 ];
			LOCAL.Return[ 4 ] = LOCAL.Values[ 1 ];

		}else if(ArrayLen( LOCAL.Values ) GTE 4){

			// Copy first four values.
			LOCAL.Return[ 1 ] = LOCAL.Values[ 1 ];
			LOCAL.Return[ 2 ] = LOCAL.Values[ 2 ];
			LOCAL.Return[ 3 ] = LOCAL.Values[ 3 ];
			LOCAL.Return[ 4 ] = LOCAL.Values[ 4 ];

		}

		// Return results
		return LOCAL.Return;
	}

	/**
	* @hint Parses the background short-hand and sets the equivalent CSS properties.
	* @PropertyMap I am the CSS hash map being updated.
	* @Value The background short hand value.
	* 
	* @output false
	*/
	public void function SetBackground(required struct PropertyMap,required string Value ){

		// Set up base properties that make up the background short hand.
		LOCAL.CSS[ "background-attachment" ] = "";
		LOCAL.CSS[ "background-color" ] = "";
		LOCAL.CSS[ "background-image" ] = "";
		LOCAL.CSS[ "background-position" ] = "";
		LOCAL.CSS[ "background-repeat" ] = "";

		// Get property tokens.
		LOCAL.Tokens = GetPropertyTokens( ARGUMENTS.Value );

		/*
			Now that we have all of our tokens, we are going to loop over the
			tokens and the properties and try to apply each. We want to apply
			tokens with the hardest to accomodate first.
		*/
		LOCAL.PropArray = ["background-attachment","background-position","background-repeat","background-image","background-color"];
		for( LOCAL.Token in LOCAL.Tokens ) {

			// Loop over properties, most restrictive first.
			for( LOCAL.Property in LOCLA.PropArray ){

				//Check to see if this value is valid. If this property
				//already has a value, then skip.

				if (
					(NOT Len( LOCAL.CSS[ LOCAL.Property ] )) AND
					IsValidValue( LOCAL.Property, LOCAL.Token )
					){

					// Assign to property.
					LOCAL.CSS[ LOCAL.Property ] = LOCAL.Token;

					// Move to next token.
					break;

				}

			}

		}


		// Loop over local CSS to apply property
		for( LOCAL.Property in LOCAL.CSS ){

		// Set properties.
			if( Len( LOCAL.CSS[ LOCAL.Property ] ) ){
				ARGUMENTS.PropertyMap[ LOCAL.Property ] = LOCAL.CSS[ LOCAL.Property ];
			}

		}

		return;
	}

	/**
	* @hint returns XSSFColor by colorName.  Since POI 4.0 you have to provide the workbench IndexedColorMap:
	* @ColorName The name of color
	* @Value The border short hand value.
	* 
	* @output false
	*/
	public any function getXSSFColorByName( string colorName="BLACK" ){

		//Check if the color exists, default to Black
		if( ! StructKeyExists(VARIABLES.poiColors, arguments.colorName ) ){
			ARGUMENTS.colorName = "BLACK";
		}

		local.color = VARIABLES.javaLoader.create( VARIABLES.classes.color ).init( VARIABLES.indexedColorMap );
		local.index = VARIABLES.IndexedColors.valueOf( JavaCast("string", ucase( ARGUMENTS.colorName ) ) ).getIndex();
		local.color.setIndexed( JavaCast("Int", local.index ) );

		return local.color;
      
	}

	/**
	* @hint Parses the border short-hand and sets the equivalent CSS properties.
	* @PropertyMap I am the CSS hash map being updated.
	* @Name The name of the pseudo property that we want to set.
	* @Value The border short hand value.
	* 
	* @output false
	*/
	public void function SetBorder(required struct PropertyMap,required string Name, required string Value ){

		//Set up base properties. We will use the top-border as our base
		//since all borders act the same and we have validation set up for it.

		LOCAL.CSS = {};
		LOCAL.CSS[ "border-top-width" ] = "";
		LOCAL.CSS[ "border-top-color" ] = "";
		LOCAL.CSS[ "border-top-style" ] = "";

		// Get property tokens.
		LOCAL.Tokens = GetPropertyTokens( ARGUMENTS.Value );

		//Now that we have all of our tokens, we are going to loop over the
		//	tokens and the properties and try to apply each. We want to apply
		//	tokens with the hardest to accomodate first.

		LOCAL.PropArray = ["border-top-style","border-top-width","border-top-color"];
		for( LOCAL.Token in LOCAL.Tokens ){

			//Loop over properties, most restrictive first.
			for( LOCAL.Property in LOCAL.PropArray ){
				//Check to see if this value is valid. If this property
				//	already has a value, then skip.
				if( (NOT Len( LOCAL.CSS[ LOCAL.Property ] )) AND
					IsValidValue( LOCAL.Property, LOCAL.Token )
					 ){
					//Assign to property.
					LOCAL.CSS[ LOCAL.Property ] = LOCAL.Token;

					//Move to next token.
					break;
				}

			}

		}

		//If we are dealing with the main border, then we have to apply
		//	these results to all four borders. Otherwise, we are only dealing
		//	with the given property.

		if (ARGUMENTS.Name EQ "border"){

			//All four borders.
			LOCAL.propertyArray = ["border-top","border-right","border-bottom","border-left"];
		} else{

			//just the given property.
			LOCAL.propertyArray = ListToArray(ARGUMENTS.Name);
		}
		for( LOCAL.Property in LOCAL.propertyArray ){
			//Loop over list to apply CSS.

			// Set properties.
			if ( Len( LOCAL.CSS[ "border-top-color" ] ) ){
				ARGUMENTS.PropertyMap[ "#LOCAL.Property#-color" ] = LOCAL.CSS[ "border-top-color" ];
			}

			if( Len( LOCAL.CSS[ "border-top-style" ] ) ){
				ARGUMENTS.PropertyMap[ "#LOCAL.Property#-style" ] = LOCAL.CSS[ "border-top-style" ];
			}

			if( Len( LOCAL.CSS[ "border-top-width" ] ) ){
				ARGUMENTS.PropertyMap[ "#LOCAL.Property#-width" ] = LOCAL.CSS[ "border-top-width" ];
			}
		}

		return;
	}


	/**
	* @hint Parses the font short-hand and sets the equivalent CSS properties.
	* @PropertyMap I am the CSS hash map being updated.
	* @Value The font short hand value.
	* 
	* @output false
	*/
	public void function SetFont(required struct PropertyMap,required string Value ){

		//Set up base properties that make up the font short hand.
		LOCAL.CSS[ "font-family" ] = "";
		LOCAL.CSS[ "font-size" ] = "";
		LOCAL.CSS[ "font-style" ] = "";
		LOCAL.CSS[ "font-weight" ] = "";

		// Get property tokens.
		LOCAL.Tokens = GetPropertyTokens( ARGUMENTS.Value );


		//Now that we have all of our tokens, we are going to loop over the
		//tokens and the properties and try to apply each. We want to apply
		//tokens with the hardest to accomodate first.


		LOCAL.fontProps = ["font-style","font-size","font-weight","font-family"];
		for( LOCAL.Token in LOCAL.Tokens ){
			//Loop over properties, most restrictive first.
			for( LOCAL.Property in LOCAL.fontProps ){
				//Check to see if this value is valid. If this property
			    //already has a value, then skip.
				if( (NOT Len( LOCAL.CSS[ LOCAL.Property ] )) AND
					IsValidValue( LOCAL.Property, LOCAL.Token )
					){
					LOCAL.CSS[ LOCAL.Property ] = LOCAL.Token;
					break;
				}
			}
		}
		for( LOCAL.Property in LOCAL.CSS ){
			//Loop over local CSS to apply property.
			if(  Len( LOCAL.CSS[ LOCAL.Property ] ) ){
			// Set properties.
				ARGUMENTS.PropertyMap[ LOCAL.Property ] = LOCAL.CSS[ LOCAL.Property ];
			}
		}

		return;
	}

	/**
	* @hint Parses the list style short-hand and sets the equivalent CSS properties.
	* @PropertyMap I am the CSS hash map being updated.
	* @Value The list style short hand value
	* 
	* @output false
	*/
	public void function SetListStyle( required struct PropertyMap,required string Value ){

		//Set up base properties that make up the list style short hand.
		LOCAL.CSS[ "list-style-image" ] = "";
		LOCAL.CSS[ "list-style-position" ] = "";
		LOCAL.CSS[ "list-style-type" ] = "";
		//Get property tokens.
		LOCAL.Tokens = GetPropertyTokens( ARGUMENTS.Value );
		// Now that we have all of our tokens, we are going to loop over the
		// tokens and the properties and try to apply each. We want to apply
		// tokens with the hardest to accomodate first.

		LOCAL.listArray = ["list-style-type","list-style-image","list-style-position"];
		for( LOCAL.Token in LOCAL.Tokens ){
			//Check to see if this value is valid. If this property
			//already has a value, then skip.
			for(LOCAL.Property in LOCAL.listArray){
				if( (NOT Len( LOCAL.CSS[ LOCAL.Property ] )) AND
					IsValidValue( LOCAL.Property, LOCAL.Token )
				   ){
					LOCAL.CSS[ LOCAL.Property ] = LOCAL.Token;
					break;
				}
			}
		}

		for( LOCAL.Property in LOCAL.CSS ){
			if( Len( LOCAL.CSS[ LOCAL.Property ] ) ){
				ARGUMENTS.PropertyMap[ LOCAL.Property ] = LOCAL.CSS[ LOCAL.Property ];
			}
		}
		return;
	}

		/**
		* @hint Parses the margin short hand and sets the equivalent properties.
		* @PropertyMap I am the CSS hash map being updated.
		* @Value The margin short hand value.
		* 
		* @output false
		*/
	public void function SetMargin(required struct PropertyMap, required string Value ){
		LOCAL.Metrics = ParseQuadMetric( ARGUMENTS.Value );
		if( IsValidValue( "margin-top", LOCAL.Metrics[ 1 ] ) ){
			ARGUMENTS.PropertyMap[ "margin-top" ] = LOCAL.Metrics[ 1 ];
		}
		if( IsValidValue( "margin-right", LOCAL.Metrics[ 2 ] ) ){
			ARGUMENTS.PropertyMap[ "margin-right" ] = LOCAL.Metrics[ 2 ];
		}
		if( IsValidValue( "margin-bottom", LOCAL.Metrics[ 3 ] ) ){
			ARGUMENTS.PropertyMap[ "margin-bottom" ] = LOCAL.Metrics[ 3 ];
		}
		if( IsValidValue( "margin-left", LOCAL.Metrics[ 4 ] ) ){
			ARGUMENTS.PropertyMap[ "margin-left" ] = LOCAL.Metrics[ 4 ];
		}
		return;

	}

	/**
	* @hint Parses the padding short hand and sets the equivalent properties.
	* @PropertyMap I am the CSS hash map being updated.
	* @Value The padding short hand value.
	* 
	* @output false
	*/
	public void function SetPadding( required struct PropertyMap, required string Value){
		LOCAL.Metrics = ParseQuadMetric( ARGUMENTS.Value );
		if( IsValidValue( "padding-top", LOCAL.Metrics[ 1 ]) ){
			ARGUMENTS.PropertyMap[ "padding-top" ] = LOCAL.Metrics[ 1 ];
		}
		if( IsValidValue( "padding-right", LOCAL.Metrics[ 2 ] ) ){
			ARGUMENTS.PropertyMap[ "padding-right" ] = LOCAL.Metrics[ 2 ];
		}
		if( IsValidValue( "padding-bottom", LOCAL.Metrics[ 3 ] ) ){
			ARGUMENTS.PropertyMap[ "padding-bottom" ] = LOCAL.Metrics[ 3 ];
		}
		if( IsValidValue( "padding-left", LOCAL.Metrics[ 4 ] ) ){
			ARGUMENTS.PropertyMap[ "padding-left" ] = LOCAL.Metrics[ 4 ];
		}
		return;
	}

	/**
	* @hint for applying ranges at the end of the sheet tag. This provides support for colSpan AND rowSpan
	* @sheet The workbook sheet being applied to
	* @col The target column ( pre adjusted for 0 based index )
	* @colspan The number of columns to span (pre adjusted for 0 based index)
	* @row The target row ( pre adjusted for 0 based index )
	* @rowSpan the number of rows (pre adjusted for 0 based index )
	* 
	* @output false
	*/
	public void function ApplyRange(required any sheet, required numeric col, required numeric colspan, required numeric row, required numeric rowspan){
		//arguments were set to 0 based index prior to call
		LOCAL.cell = ARGUMENTS.sheet.getRow( ARGUMENTS.row ).getCell( ARGUMENTS.col );
		LOCAL.range = VARIABLES.javaLoader.create(  VARIABLES.classes.cellRangeAddress ).init(JavaCast("int",ARGUMENTS.row), JavaCast("int",ARGUMENTS.rowSpan), JavaCast("int",ARGUMENTS.col), JavaCast("int",ARGUMENTS.colSpan) );
		LOCAL.cellStyle = LOCAL.cell.getCellStyle();
		ARGUMENTS.sheet.addMergedRegion( LOCAL.range );

		if( StructKeyExists(LOCAL,"cellStyle") ){

			VARIABLES.RegionUtil.setBorderBottom( LOCAL.cellStyle.getBorderBottom(),  LOCAL.range, ARGUMENTS.sheet );
	    	VARIABLES.RegionUtil.setBorderTop(    LOCAL.cellStyle.getBorderTop(),     LOCAL.range, ARGUMENTS.sheet );
	    	VARIABLES.RegionUtil.setBorderLeft(   LOCAL.cellStyle.getBorderLeft(),    LOCAL.range, ARGUMENTS.sheet );
	    	VARIABLES.RegionUtil.setBorderRight(  LOCAL.cellStyle.getBorderRight(),   LOCAL.range, ARGUMENTS.sheet );

	    	VARIABLES.RegionUtil.setBottomBorderColor( LOCAL.cellStyle.getBottomBorderColor(),  LOCAL.range, ARGUMENTS.sheet );
	    	VARIABLES.RegionUtil.setTopBorderColor(    LOCAL.cellStyle.getTopBorderColor(),     LOCAL.range, ARGUMENTS.sheet );
	    	VARIABLES.RegionUtil.setLeftBorderColor(   LOCAL.cellStyle.getLeftBorderColor(),    LOCAL.range, ARGUMENTS.sheet );
	    	VARIABLES.RegionUtil.setRightBorderColor(  LOCAL.cellStyle.getRightBorderColor(),   LOCAL.range, ARGUMENTS.sheet );
    	}  	
	}
	

}