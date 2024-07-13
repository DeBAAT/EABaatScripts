//[group=BaatScriptLib]
!INC EAScriptLib.JavaScript-Logging

/**
 * @file JavaScript-EXCEL
 * This script library contains helper functions to assist with EXCEL Import and Export of
 * Enterprise Architect elements. 
 * 
 * Functions in this library are split into three parts: Workbooks, Import and Export.
 * Functions that assist with EXCEL Workbooks are prefixed with EXCELW,
 * EXCEL Import are prefixed with EXCELI, whereas functions that assist with EXCEL Export are prefixed 
 * with EXCELE.
 *
 * EXCEL Import can be performed by calling the function EXCELIImportFile(). EXCELIImportFile() requires
 * that the function OnExcelRowImported() be defined in the user's script to be used as a callback
 * whenever row data is read from the EXCEL file. The user defined OnExcelRowImported() can query for 
 * information about the current row through the functions EXCELIContainsColumn(), 
 * EXCELIGetColumnValueByName() and EXCELIGetColumnValueByNumber().
 *
 * To perform an EXCEL export, the user must firstly call EXCELEExportInitialize() which starts an export 
 * session. The call to EXCELEExportInitialize() specifies the file name to export to, and the set of 
 * columns that will be exported. Once the session has been initialized with a call to 
 * EXCELEExportInitialize(), the user may continually call EXCELEExportRow() to export a row to file. 
 * Once all rows have been added, the export session is closed by calling EXCELEExportFinalize(). 
 *
 * @author J. de Baat, based on JavaScript - CSV by Sparx Systems
 * @date 2024-07-13
 */

////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////
////																							////
////											EXCEL WORKBOOKS									////
////																							////
////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////
var objExcelApplication = null;   // The global variable holding the Excel application running
var objExcelWorkBook    = null;   // The global variable holding the Excel WorkBook  in the file
var objExcelWorkSheet   = null;   // The global variable holding the Excel WorkSheet in the WorkBook

const EXCEL_DELIMITER   = ",";
const TAG_PREFIX        = "TAG_";

/**
 * Returns the opened EXCEL Workbook from the fileName provided.
 *
 * @param[in] fileName (string) The path to the EXCEL file to open.
*/
function EXCELWOpenWorkbook( fileName /* : String */ ) /* : Microsoft.Office.Interop.Excel._Workbook */
{

	objExcelWorkBook = null;

	// Get objExcelWorkBook from the fileName provided
	try {
		objExcelWorkBook = objExcelApplication.Workbooks.Open( fileName );
	} catch (err) {
		LOGError( "EXCELWOpenWorkbook catched error " + err.message + "!" );
		objExcelWorkBook = null;
	}
	if ( objExcelWorkBook != null ) {
		// objExcelWorkBook FOUND
		Session.Output("EXCELWOpenWorkbook Found Workbook with fileName = " + fileName + " !" );
	} else {
		// Make sure we have a file
		// Session.Output( "EXCELWGetFileName Started with strExcelFileName = " + strExcelFileName + " !" );
		strExcelFileName = EXCELWCheckExcelFileName( strExcelFileName );
		Session.Output( "EXCELWGetFileName Checked strExcelFileName = " + strExcelFileName + " !" );

		// If the file does not exist, then create one
		try {
			objExcelWorkBook = objExcelApplication.Workbooks.Add();
			// Session.Output("EXCELWOpenWorkbook objExcelApplication contains " + objExcelApplication.Workbooks.Count + " Workbooks after Add!" );
			Session.Output("EXCELWOpenWorkbook created a new Workbook with fileName = " + fileName + " !" );
		} catch (err) {
			LOGError( "EXCELWOpenWorkbook could NOT open nor add fileName " + fileName + ", catched error " + err.message + "!" );
			objExcelWorkBook = null;
		}
	}

	return objExcelWorkBook;

}

/**
 * Check theExcelFileName and create new file when the file does not exist
 */
function EXCELWCheckExcelFileName( theExcelFileName ) /* : string */
{

	let strExcelFileName = "";
	strExcelFileName     = theExcelFileName;

	// Make sure we have a filename
	if ( ( strExcelFileName == null ) || ( strExcelFileName.length == 0 ) ) {
		LOGError( "EXCELWCheckExcelFileName could NOT check empty strExcelFileName!" );
		return null;
	}

	// Session.Output( "EXCELWCheckExcelFileName checking strExcelFileName = " + strExcelFileName + "!!!" );

	// Create a fileSystemObject to check the existence of the file chosen
	try {
		var fileSystemObject = new COMObject( "Scripting.FileSystemObject" );
	} catch (err) {
		LOGError( "EXCELWCheckExcelFileName catched error " + err.message + "!" );
		return null;
	}
	if ( fileSystemObject.FileExists( strExcelFileName ) ) {

		// File exists so return the name found
		// Session.Output( "EXCELWCheckExcelFileName found strExcelFileName = " + strExcelFileName + " in filesystem!!!" );
		fileSystemObject = null;
		return strExcelFileName;
	} else {

		// If the file does not exist, then create one
		try {
			objExcelWorkBook = objExcelApplication.Workbooks.Add();
		} catch (err) {
			LOGError( "EXCELWCheckExcelFileName catched error " + err.message + "!" );
			objExcelWorkBook = null;
		}
		// Session.Output("EXCELWCheckExcelFileName objExcelApplication DefaultFilePath = " + objExcelApplication.DefaultFilePath + "!!!" );
		// Session.Output("EXCELWCheckExcelFileName objExcelApplication contains " + objExcelApplication.Workbooks.Count + " Workbooks after Add!" );
		if ( objExcelWorkBook != null ) {
			let strExcelWorkBookFullName = "";

			// If the file added, then save it
			objExcelWorkBook.Save( true );
			strExcelWorkBookFullName = objExcelWorkBook.FullName;
			// Session.Output("EXCELWCheckExcelFileName saved objExcelWorkBook with FullName = " + objExcelWorkBook.FullName + " after Add!" );

			// Close the newly added objExcelWorkBook to move the file
			objExcelWorkBook.Close();
			objExcelWorkBook = null;

			// Check filesystem whether file exists
			let objExcelFile = null;
			try {
				objExcelFile = fileSystemObject.GetFile( strExcelWorkBookFullName );
				objExcelFile.Move( strExcelFileName );
				objExcelFile = null;
			} catch (err) {
				LOGError( "EXCELWCheckExcelFileName could NOT open objExcelFile for strExcelFileName " + strExcelWorkBookFullName + "!" );
				LOGError( "EXCELWCheckExcelFileName catched error " + err.message + "!" );
				fileSystemObject = null;
				return null;
			}

			// Session.Output("EXCELWCheckExcelFileName created a new Workbook with strExcelFileName = " + strExcelFileName + " !" );
		} else {
			LOGError( "EXCELWCheckExcelFileName could NOT open nor add strExcelFileName " + strExcelFileName + "!" );
			fileSystemObject = null;
			return null;
		}
	}

	fileSystemObject = null;
	return strExcelFileName;
}

/**
 * Saves the provided EXCEL Workbook with the fileName provided.
 *
 * @param[in] fileName (string) The path to the EXCEL file to save to.
*/
function EXCELWSaveWorkbook( fileName /* : String */ ) /* : void */
{
	if ( objExcelWorkBook != null ) {
		try {
			// objExcelWorkBook FOUND
			// Session.Output("EXCELWSaveWorkbook saves Workbook with fileName = " + fileName + " !" );
			objExcelWorkBook.SaveAs( fileName );
			// Session.Output("EXCELWSaveWorkbook objExcelWorkBook FullName = " + objExcelWorkBook.FullName + " after SaveAs!" );
			objExcelWorkBook.Save(  true  );
			// Session.Output("EXCELWSaveWorkbook objExcelWorkBook FullName = " + objExcelWorkBook.FullName + " after Save!" );
			// Session.Output("EXCELWSaveWorkbook objExcelApplication contains " + objExcelApplication.Workbooks.Count + " Workbooks!" );
			// Session.Output("EXCELWSaveWorkbook objExcelWorkBook contains " + objExcelWorkBook.Sheets.Count + " Sheets!" );
			// Session.Output("EXCELWSaveWorkbook saved Workbook with fileName = " + fileName + " !" );
		} catch (err) {
			LOGError( "EXCELWSaveWorkbook catched error " + err.message + "!" );
		}
	} else {
		LOGError( "EXCELWSaveWorkbook could NOT save Workbook to fileName " + fileName + "!" );
	}

}

/**
 * Closes the provided EXCEL Workbooks opened before.
 *
 * @param[in] saveWorkbook (boolean) If set to true, the Workbook is saved without asking for confirmation
*/
function EXCELWCloseWorkbooks( saveWorkbook /* : boolean */ ) /* : void */
{
	try {
		if ( ( saveWorkbook != null ) && ( saveWorkbook ) ) {
			objExcelWorkBook.Save( true );
		}
		objExcelApplication.Workbooks.Close();
	} catch (err) {
		LOGError( "EXCELWCloseWorkbooks catched error " + err.message + "!" );
	}

}

/**
 * Start the EXCEL Application.
 *
*/
function EXCELWStartExcelApplication() /* : void */
{

	try {
		objExcelApplication = new COMObject( "Excel.Application", true );
		if ( objExcelApplication != null ) {
			// objExcelApplication STARTED
			// Session.Output("EXCELWStartExcelApplication started Excel.Application !" );
		} else {
			LOGError( "EXCELWStartExcelApplication could NOT start Excel.Application!" );
		}
	} catch (err) {
		LOGError( "EXCELWStartExcelApplication catched error " + err.message + "!" );
		objExcelApplication = null;
	}

	return objExcelApplication;

}

/**
 * Stop the provided EXCEL Application started before.
 *
*/
function EXCELWStopExcelApplication() /* : void */
{

	try {
		objExcelApplication.Quit();
	} catch (err) {
		LOGError( "EXCELWStopExcelApplication catched error " + err.message + "!" );
		objExcelApplication = null;
	}

}

/**
 * Gets the Worksheet as part of the provided EXCEL Workbooks using the sheetName provided.
 * If the Worksheet does not exist and addSheet is true, it is added to the Workbooks Collection
 *
 * @param[in] sheetName (string) The name of the Worksheet in the EXCEL file to get or create.
 * @param[in] addSheet (boolean) If set to true, the Worksheet is added if it does not exist
*/
function EXCELWGetWorksheet( sheetName /* : String */, addSheet /* : boolean */ ) /* : void */
{

	var curExcelWorkSheet = null;

	// Check valid objExcelWorkBook
	if ( objExcelWorkBook == null ) {
		// objExcelWorkBook NOT FOUND
		LOGError( "EXCELWGetWorksheet Could NOT get Worksheet " + sheetName + " because objExcelWorkBook NOT opened!" );
		return curExcelWorkSheet;
	}

	// Get the curExcelWorkSheet
	try {
		curExcelWorkSheet = objExcelWorkBook.Sheets.Item( sheetName );
	} catch (err) {
		LOGError( "EXCELWGetWorksheet catched error " + err.message + "!" );
		curExcelWorkSheet = null;
	}
	if ( curExcelWorkSheet == null ) {
		// Create the curExcelWorkSheet if it is not found and addSheet is true
		if ( addSheet ) {
			try {
				curExcelWorkSheet = objExcelWorkBook.Sheets.Add();
				curExcelWorkSheet.Name = sheetName;
				// Session.Output("EXCELWGetWorksheet created Worksheet with sheetName = " + curExcelWorkSheet.Name + " !" );
			} catch (err) {
				LOGError( "EXCELWGetWorksheet Could NOT get Worksheet with sheetName = " + sheetName + " !" );
				LOGError( "EXCELWGetWorksheet catched error " + err.message + "!" );
				curExcelWorkSheet = null;
			}
		}
	}

	return curExcelWorkSheet;

}

/**
 * Get the strExcelFileName to process
 *
 * @param[in] fileName (string) The path to the EXCEL file to open as specified by default
 * @param[in] readOnly (boolean) If set to true, the file is opened for read otherwise for write
 */
function EXCELWGetFileName( fileName /* : String */, readOnly /* : boolean */ ) /* : string */
{
	// Define some variables and values
	var projectInterface as EA.Project
	var strExcelFileName, FilterString, Filterindex, Flags, InitialDirectory, OpenorSave;
	strExcelFileName = fileName;
	FilterString     = "Excel Files (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls|All Files (*.*)|*.*||";
	Filterindex      = 1;
	Flags            = 0;
	InitialDirectory = "";
	if ( readOnly ) {
		OpenorSave   = 0;
	} else {
		OpenorSave   = 1;
	}

	// Get the strExcelFileName to get the information from
	try {
		projectInterface = Repository.GetProjectInterface();
		strExcelFileName = projectInterface.GetFileNameDialog( strExcelFileName, FilterString, Filterindex, Flags, InitialDirectory, OpenorSave );
	} catch (err) {
		LOGError( "EXCELWGetFileName catched error " + err.message + "!" );
		strExcelFileName = "";
	}

	// Make sure we have a filename
	if ( strExcelFileName.length == 0 ) {
		LOGError( "EXCELWGetFileName Could NOT get a valid fileName, starting with : " + fileName + " !" );
		return null;
	} else {
		// Make sure we have a file
		// Session.Output( "EXCELWGetFileName Found strExcelFileName = " + strExcelFileName + " !" );
		strExcelFileName = EXCELWCheckExcelFileName( strExcelFileName );
		Session.Output( "EXCELWGetFileName Checked strExcelFileName = " + strExcelFileName + " !" );
	}

	return strExcelFileName;
}

/**
 * If theValue starts with thePrefix then return value without prefix else return empty string
 *
 * @param[in] theValue (string) The value to test.
 * @param[in] thePrefix (string) The prefix to test against.
 */
function EXCELGGetValueWithoutPrefix( theValue, thePrefix )
{

	try {
		const thePrefixLength = thePrefix.length;
		// Session.Output("EXCELGGetValueWithoutPrefix started with thePrefix : " + thePrefix + ", length = " + thePrefixLength + " !" );

		//	If theValue starts with thePrefix then return value without prefix
		if ( theValue.substring( 0, thePrefixLength ).toLowerCase() == thePrefix.toLowerCase() ) {
			// Session.Output( "EXCELGGetValueWithoutPrefix found theValue.substring(" + thePrefixLength + ")= " + theValue.substring( thePrefixLength ) + " !" );
			return theValue.substring( thePrefixLength );
		}
	} catch (err) {
		LOGError( "EXCELGGetValueWithoutPrefix catched error " + err.message + "!" );
	}

	return "";

}

////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////
////																							////
////											EXCEL IMPORT									////
////																							////
////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////
var excelImportColumnMap;				// : Map with column index values related to column headers
var excelImportColumnTagsMap;			// : Map with tag names found within the column headers
var excelImportColumnList;				// : Array with column headers
var excelImportCurrentRow;				// : Array with values found in this row
var excelImportIsImporting = false;

/**
 * Imports the provided EXCEL Workbooks as indicated by the sheetName.
 * The user function OnExcelRowImported() is called to notify the user
 * script that row data is available. The user script may then call EXCELIContainsColumn() to see
 * if the current row contains a particular named column, and EXCELIGetColumnValueByName() or
 * EXCELIGetColumnValueByNumber() to obtain the field value of a particular column.
 *
 * @param[in] sheetName (string) The name of the Worksheet in the EXCEL file to import.
 * 
 * @param[in] firstRowContainsHeadings (boolean) If set to true, the values of the first row will be parsed
 * as column headings. EXCELIContainsColumn() and EXCELIGetColumnValueByName() will only work if 
 * firstRowContainsHeadings is set to true.
 */
function EXCELIImportSheet( sheetName /* : String */, firstRowContainsHeadings /* : boolean */ ) /* : void */
{
	if ( !excelImportIsImporting ) {

		// Check valid objExcelWorkBook
		if ( objExcelWorkBook == null ) {
			// objExcelWorkBook NOT FOUND
			LOGError( "EXCELIImportSheet Could NOT get Worksheet " + sheetName + " because objExcelWorkBook NOT opened!" );
			return null;
		}

		excelImportIsImporting = true;

		// Get and check the curExcelWorkSheet for sheetName
		var curExcelWorkSheet  = EXCELWGetWorksheet( sheetName, false );
		if ( curExcelWorkSheet != null ) {

			// Set up row/column caching
			let curValue    = new String( "" );
			let RowCount    = 0;
			let ColumnCount = 0;
			try {
				RowCount    = curExcelWorkSheet.UsedRange.Rows.Count;
				ColumnCount = curExcelWorkSheet.UsedRange.Columns.Count;

				excelImportColumnMap     = new Map();
				excelImportColumnTagsMap = new Map();
				excelImportColumnList    = [];
			} catch (err) {
				LOGError( "EXCELIImportSheet catched error " + err.message + "!" );
				RowCount    = 0;
				ColumnCount = 0;
			}

			// Process the Worksheet a row at a time
			for ( var curRow = 1 ; curRow <= RowCount ; curRow++ ) {

				try {

					// Reset excelImportCurrentRow for next row
					excelImportCurrentRow    = [];

					if ( curRow == 1 && firstRowContainsHeadings )
					{
						// Cache column heading positions
						for ( var curCol = 1 ; curCol <= ColumnCount ; curCol++ ) {
							curValue.text = curExcelWorkSheet.Cells.Item( curRow, curCol ).Value;
							excelImportColumnMap.set( curValue.text, curCol - 1 ); // Col starts with 1, List starts with 0
							excelImportColumnList.push( curValue.text );
							// Session.Output( "EXCELIImportSheet excelImportColumnMap.set(" + curValue.text + "," + curCol + ")!!!" );
							// Session.Output( "EXCELIImportSheet excelImportColumnList.push(" + curValue.text + ")!!!" );

							if ( curValue.text != null ) {
								let importColumnTag = EXCELGGetValueWithoutPrefix( curValue.text, TAG_PREFIX );
								if ( importColumnTag.length > 0 ) {
									excelImportColumnTagsMap.set( curValue.text, importColumnTag );
									// Session.Output( "EXCELIImportSheet excelImportColumnTagsMap.set(" + curValue.text + "," + importColumnTag + ")!!!" );
								}
							}
						}
					}
					else
					{
						// Hold a reference to the current row data
						// Cache column heading positions
						for ( var curCol = 1 ; curCol <= ColumnCount ; curCol++ ) {
							excelImportCurrentRow.push( curExcelWorkSheet.Cells.Item( curRow, curCol ).Value );
							// Session.Output( "EXCELIImportSheet excelImportCurrentRow.push(" + (curCol - 1) + ") Cell(" + curRow + "," + curCol + "): " + curExcelWorkSheet.Cells.Item( curRow, curCol ).Value + "!!!" );
						}
						
						// Invoke the user script callback
						OnExcelRowImported( curRow );
					}
				} catch (err) {
					LOGError( "EXCELIImportSheet for row[" + curRow + "] catched error " + err.message + "!" );
				}

			}
		} else {
			LOGError( "EXCELIImportSheet did NOT find sheetName " + sheetName + "!"  );
		}
		
		// Clean up
		excelImportColumnMap      = null;
		excelImportColumnTagsMap  = null;
		excelImportColumnList     = null;
		excelImportCurrentRow     = null;
		excelImportIsImporting    = false;
	}
	else
	{
		LOGWarning( "Reentrant call made to EXCELIImportSheet(). EXCELIImportSheet() should not be called from within OnExcelRowImported()!" );
	}
}

/**
 * Advises whether the current import row contains a field value for the specified column name.
 *
 * NOTE: The function only works if EXCELImportFile() was called with the firstRowContainsHeadings
 * paremeter set to true.
 *
 * @param[in] (String) The name of the column to check for
 *
 * @return A boolean indicating whether the current import row contains a field value for the 
 * specified column name.
 */
function EXCELIContainsColumn( columnName /* : String */ ) /* : boolean */
{
	var result = false;
	
	if ( excelImportIsImporting )
	{
		// Get the column number of the specified named column
		var columnNumber = __EXCELIGetColumnNumber( columnName );
	
		// If the column is in range then it exists!
		if ( columnNumber >= 0 && excelImportCurrentRow != null && columnNumber < excelImportCurrentRow.length ) {
			result = true;
		}
	}
	else
	{
		LOGWarning( "No import currently running. EXCELIContainsColumn() should only be called from within OnExcelRowImported()" );
	}
	
	return result;
}

/**
 * Returns the value of the field in the current import row with the specified column name
 *
 * NOTE: The function only works if EXCELImportFile() was called with the firstRowContainsHeadings
 * paremeter set to true.
 *
 * @param[in] columnName (String) The name of the column whose value will be retrieved.
 *
 * @return The current import row's field value for the specified column
 */
function EXCELIGetColumnValueByName( columnName /* : String */ ) /* : variant */
{	
	var result;
	
	if ( excelImportIsImporting )
	{
		var columnNumber = __EXCELIGetColumnNumber( columnName );
		result = EXCELIGetColumnValueByNumber( columnNumber );
	}
	else
	{
		LOGWarning( "No import currently running. EXCELIGetColumnValueByName() should only be called from within OnExcelRowImported()" );
	}
	
	return result;
}

/**
 * Returns the value of the field in the current import row with the specified column number
 *
 * @param[in] columnNumber (number) The index of the column whose value will be retrieved.
 *
 * @return The current import row's field value for the specified column
 */
function EXCELIGetColumnValueByNumber( columnNumber /* : number */ ) /* : variant */
{
	var result;
	
	if ( excelImportIsImporting )
	{
		if ( columnNumber >= 0 && excelImportCurrentRow != null && columnNumber < excelImportCurrentRow.length ) {
			result = excelImportCurrentRow[columnNumber];
		}
	}
	else
	{
		LOGWarning( "No import currently running. EXCELIGetColumnValueByNumber() should only be called from within OnExcelRowImported()" );
	}
		
	return result;
}

/**
 * Returns all column names that are not considered standard as an array of Strings.
 */
function EXCELIGetNonStandardElementColumns() /* : Array */
{
	var result = [];	
	var standardColumns = new String(";Abstract;Alias;Author;ClassifierName;Complexity;Created");
	standardColumns += ";Difficulty;GenFile;GenType;Header1;Header2;IsActive;IsLeaf";
	standardColumns += ";IsNew;IsSpec;Locked;Multiplicity;Name;Notes;Persistence;Phase;Priority";
	standardColumns += ";RunState;Status;Stereotype;Subtype;Tablespace;Tag;TreePos;Type;Version";
	standardColumns += ";Visibility;";
	
	if ( excelImportIsImporting )
	{
		for ( var i = 0 ; i < excelImportCurrentRow.length ; i++ )
		{
			if ( i < excelImportColumnList.length )
			{
				var columnName = excelImportColumnList[i];
				if ( standardColumns.indexOf( ";" + columnName + ";" ) == -1 )
				{
					result.push( columnName );
				}
			}
			
		}
	}
	else
	{
		LOGWarning( "No import currently running. EXCELIGetNonStandardElementColumns() should only be called from within OnExcelRowImported()" );
	}
	
	return result;
}

/**
 * Sets the properties on the specified Element if there is a corresponding value for them in the 
 * current row.
 *
 * Element properties that are not set by this function include:
 * 	- Read only properties
 * 	- Collection properties
 * 	- Properties that contain relational information (eg IDs/GUIDs of other elements, connectors 
 *	or packages.
 *	- Modified Date property (this is property is automatically overwritten by the automation 
 *	interface when the element is saved)
 *	- Properties that are themselves a comma separated list
 *
 * @param[in] elementForRow (EA.Element) The element whose properties will be set with the current row's 
 * values
 */
function EXCELISetStandardElementFieldValues( elementForRow /* : EA.Element */ ) /* : void */
{
	if ( excelImportIsImporting )
	{
		var theElement as EA.Element;
		theElement = elementForRow;
		
		if ( theElement != null )
		{
			if ( EXCELIContainsColumn("Abstract") )
				theElement.Abstract = EXCELIGetColumnValueByName("Abstract");
			
			// ActionFlags - Not included (Comma Separated)
			
			if ( EXCELIContainsColumn("Alias") )
				theElement.Alias = EXCELIGetColumnValueByName("Alias");
			
			// Attributes - Not included (Collection)
			// AttributesEx - Not included (Collection)
			
			if ( EXCELIContainsColumn("Author") )
				theElement.Author = EXCELIGetColumnValueByName("Author");
			
			// BaseClasses - Not included (Collection)
			// ClassifierID - Not included (Relational)
			
			if ( EXCELIContainsColumn("ClassifierName") )
				theElement.ClassifierName = EXCELIGetColumnValueByName("ClassifierName");
			
			// ClassifierType - Not included (Read Only)
			
			if ( EXCELIContainsColumn("Complexity") )
				theElement.Complexity = EXCELIGetColumnValueByName("Complexity");
			
			// CompositeDiagram - Not included (Relational)
			// Connectors - Not included (Collection)
			// Constraints - Not included (Collection)
			// ConstraintsEx - Not included (Collection)
					
			if ( EXCELIContainsColumn("Created") )
			{
				var dateString = EXCELIGetColumnValueByName("Created");
				var asEADate = DTParseEADate( dateString );
				theElement.Created = asEADate;
			}
			
			// CustomProperties - Not included (Collection)
			// Diagrams - Not included (Collection)
			
			if ( EXCELIContainsColumn("Difficulty") )
				theElement.Difficulty = EXCELIGetColumnValueByName("Difficulty");
			
			// Efforts - Not included (Collection)
			// ElementGUID - Not included (Read Only)
			// ElementID - Not included (Read Only)
			// Elements - Not included (Collection)
			// EmbeddedElements - Not included (Read Only)
			// EventFlags - Not included (Comma Separated)	
			// ExtensionPoints - Not included (Comma Separated)
			// Files - Not included (Collection)
			
			if ( EXCELIContainsColumn("GenFile") )
				theElement.Genfile = EXCELIGetColumnValueByName("GenFile");
			
			// Genlinks - Not included (Relational)
			
			if ( EXCELIContainsColumn("GenType") )
				theElement.Gentype = EXCELIGetColumnValueByName("GenType");
			
			if ( EXCELIContainsColumn("Header1") )
				theElement.Header1 = EXCELIGetColumnValueByName("Header1");
			
			if ( EXCELIContainsColumn("Header2") )
				theElement.Header1 = EXCELIGetColumnValueByName("Header2");
			
			if ( EXCELIContainsColumn("IsActive") )
				theElement.IsActive = EXCELIGetColumnValueByName("IsActive");
			
			if ( EXCELIContainsColumn("IsLeaf") )
				theElement.IsLeaf = EXCELIGetColumnValueByName("IsLeaf");
			
			if ( EXCELIContainsColumn("IsNew") )
				theElement.IsNew = EXCELIGetColumnValueByName("IsNew");
			
			if ( EXCELIContainsColumn("IsSpec") )
				theElement.IsSpec = EXCELIGetColumnValueByName("IsSpec");
			
			// Issues - Not included (Collection)
			
			if ( EXCELIContainsColumn("Locked") )
				theElement.Locked = EXCELIGetColumnValueByName("Locked");
			
			// MetaType - Not included (Read Only)
			// Methods - Not included (Collection)
			// Metrics - Not included (Collection)
			// MiscData - Not included (Read Only)
			// Modified - Not included (Overwritten)
			
			if ( EXCELIContainsColumn("Multiplicity") )
				theElement.Multiplicity = EXCELIGetColumnValueByName("Multiplicity");
			
			if ( EXCELIContainsColumn("Name") )
			{
				theElement.Name = EXCELIGetColumnValueByName("Name");
			}
			
			if ( EXCELIContainsColumn("Notes") )
				theElement.Notes = EXCELIGetColumnValueByName("Notes");
			
			// ObjectType - Not included (Read Only)
			// PackageID - Not included (Relational)
			// ParentID - Not included (Relational)
			// Partitions - Not included (Collection)
			
			if ( EXCELIContainsColumn("Persistence") )
				theElement.Persistence = EXCELIGetColumnValueByName("Persistence");
			
			if ( EXCELIContainsColumn("Phase") )
				theElement.Persistence = EXCELIGetColumnValueByName("Phase");
			
			if ( EXCELIContainsColumn("Priority") )
				theElement.Persistence = EXCELIGetColumnValueByName("Priority");
			
			// Properties - Not included (Collection)
			// PropertyType - Not included (Relational)
			// Realizes - Not included (Collection)
			// Requirements - Not included (Collection)
			// RequirementsEx - Not included (Collection)
			// Resources - Not included (Collection)
			// Risks - Not included (Collection)
			
			if ( EXCELIContainsColumn("RunState") )
				theElement.Persistence = EXCELIGetColumnValueByName("RunState");
			
			// Scenarios - Not included (Collection)
			// State Transitions - Not included (Collection)
			
			if ( EXCELIContainsColumn("Status") )
				theElement.Status = EXCELIGetColumnValueByName("Status");
			
			if ( EXCELIContainsColumn("Stereotype") )
				theElement.Stereotype = EXCELIGetColumnValueByName("Stereotype");
			
			// StereotypeEx - Not included (Comma Separated)
			
			if ( EXCELIContainsColumn("Subtype") )
				theElement.Subtype = EXCELIGetColumnValueByName("Subtype");
			
			if ( EXCELIContainsColumn("Tablespace") )
				theElement.Tablespace = EXCELIGetColumnValueByName("Tablespace");
			
			if ( EXCELIContainsColumn("Tag") )
				theElement.Tag = EXCELIGetColumnValueByName("Tag");
			
			// TaggedValues - Not included (Collection)
			// TaggedValuesEx - Not included (Collection)
			// Tests - Not included (Collection)
			
			if (EXCELIContainsColumn("TreePos") )
				theElement.TreePos = EXCELIGetColumnValueByName("TreePos");
			
			if ( EXCELIContainsColumn("Type") )
				theElement.Type = EXCELIGetColumnValueByName("Type");
			
			if ( EXCELIContainsColumn("Version") )
			{
				theElement.Version = EXCELIGetColumnValueByName("Version");
			}
			
			if ( EXCELIContainsColumn("Visibility") )
			{
				theElement.Visibility = EXCELIGetColumnValueByName("Visibility");
			}
		}
	}
	else
	{
		LOGWarning( "No import currently running. EXCELISetStandardElementFieldValues() should only be called from within OnExcelRowImported()" );		
	}
}

/**
 * @private
 * Returns the index of the column with the specified name
 *
 * @param[in] columnName (String) The name of the column whose index will be retrieved
 *
 * @return The index of the column with the specified name
 */
function __EXCELIGetColumnNumber( columnName /* : String */ ) /* : number */
{
	var result = -1;
	if ( excelImportColumnMap != null && excelImportColumnMap.has(columnName) ) {
		result = excelImportColumnMap.get(columnName);
	}
	
	return result;
}

////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////
////																							////
////											EXCEL EXPORT									////
////																							////
////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////
let exportRow            = 1;		// : Integer with the value for the current Row to export to
let exportHeaderRow      = 0;		// : Integer with the value for the header Row to export to
let exportColumns        = null;	// : Array with column headers to export
let exportExcelWorkSheet = null;
let exportIsExporting    = false;

/**
 * Initialises a EXCEL Export session. This must be called before calls to EXCELEExportRow() are made. 
 * Once all rows have been exported, a corresponding call to EXCELEExportFinalize() should be made.
 *
 * @param[in] sheetName (string) The name of the Worksheet in the EXCEL file to export the values to.
 * @param[in] columns (Array) an array of column names that will be exported
 * @param[in] exportColumnHeadings (boolean) Specifies whether the first row should contain the column 
 * headings
 */
function EXCELEExportInitialize( sheetName /* : String */, columns /* : Array */, 
	exportColumnHeadings /* : boolean */ ) /* : void */
{

	if ( !exportIsExporting )
	{

		// Check if Worksheet successfully opened for writing
		if ( objExcelWorkBook == null ) {
			// objExcelWorkBook NOT FOUND
			LOGError( "EXCELEExportInitialize Could NOT get Worksheet " + sheetName + ", because objExcelWorkBook NOT opened!" );
			return null;
		}

		// Switch into exporting mode
		exportIsExporting = true;

		// Setup column array
		exportColumns = columns;

		// Get and check the exportExcelWorkSheet for sheetName
		exportExcelWorkSheet  = EXCELWGetWorksheet( sheetName, true );
		if ( exportExcelWorkSheet != null ) {

			try {
				// Clear the exportExcelWorkSheet before exporting the new values
				let RowCount    = exportExcelWorkSheet.UsedRange.Rows.Count;
				let ColumnCount = exportExcelWorkSheet.UsedRange.Columns.Count;
				// Session.Output( "EXCELEExportInitialize exportExcelWorkSheet.UsedRange to clear: (" + RowCount + "," + ColumnCount + ")!!!" );
				exportExcelWorkSheet.UsedRange.Clear();

				// Check if headers should be exported
				if ( exportColumnHeadings )
				{
					// Get the header information from the current exportRow
					exportHeaderRow = exportRow;
					for ( var curCol = 1 ; curCol <= exportColumns.length ; curCol++ ) {
						exportExcelWorkSheet.Cells.Item( exportHeaderRow, curCol ).Value = exportColumns[ curCol - 1 ];
					}
					exportRow++;
				}
			} catch (err) {
				LOGError( "EXCELEExportInitialize catched error " + err.message + "!" );
				return null;
			}
		} else {
			LOGError( "EXCELEExportInitialize did NOT find sheetName " + sheetName + "!"  );
			return null;
		}

	}
	else
	{
		LOGWarning( "EXCELEExportInitialize: EXCEL Export is already in progress" );
		return null;
	}

	// Session.Output( "EXCELEExportInitialize exportRow = " + exportRow + "!!!" );

	return exportExcelWorkSheet;

}

/**
 * Add new columns to the columns already defined.
 * Check for duplicates.
 */
function EXCELEAddExportColumns( columns /* : Array */ ) /* : Void */
{
	// Check whether exportIsExporting and exportHeaderRow used
	if ( ( exportIsExporting ) && ( exportHeaderRow > 0 ) )
	{
		try {

			let columnArray       = [];
			let exportColumnToAdd = exportColumns.length;

			// Setup column array
			columnArray = columns;

			// Extend row/column caching
			for ( var curCol = 1 ; curCol <= columnArray.length ; curCol++ ) {
				let exportColumnText = columnArray[ curCol - 1 ];
				let duplicateFound   = false;
				// Check current exportColumns to prevent duplicates
				for ( var curExCol = 0 ; curExCol < exportColumns.length ; curExCol++ ) {
					if ( exportColumns[ curExCol ] == exportColumnText ) {
						duplicateFound = true;
					}
				}
				// Only add new column when not duplicateFound
				if ( ! duplicateFound ) {
					exportColumnToAdd++;
					exportColumns.push( exportColumnText );
					exportExcelWorkSheet.Cells.Item( exportHeaderRow, exportColumnToAdd ).Value = exportColumnText;
				}
			}
		} catch (err) {
			LOGError( "EXCELEAddExportColumns catched error " + err.message + "!" );
		}
	}
	else
	{
		LOGWarning( "EXCELEAddExportColumns: EXCEL Export is not currently in progress" );
	}
}

/**
 * Finalizes an EXCEL Export session, closing file system resources required for the export. After this
 * function has been executed, further calls to EXCELEExportRow() will fail until another EXCEL Export
 * session is initialized via EXCELEExportInitialize()
 */
function EXCELEExportFinalize() /* : void */
{
	if ( exportIsExporting )
	{
		// Clean up column array
		exportColumns = null;

		// Switch out of exporting mode
		exportIsExporting = false;
	}
	else
	{
		LOGWarning( "EXCELEExportFinalize: EXCEL Export is not currently in progress" );
	}
}

/**
 * Exports a row to the EXCEL file. The valueMap parameter is used to lookup field values for the
 * columns specified when EXCELEExportInitialize() was called. Values in valueMap that do not 
 * correspond to a valid column will not be exported.
 *
 * @param[in] valueMap (Map) A Map of field values where key=Column Name, value=Field Value
 */

function EXCELEExportRow( valueMap /* : Map */ ) /* : void */
{
	if ( exportIsExporting )
	{

		// Check if Worksheet successfully opened for writing
		if ( exportExcelWorkSheet == null ) {
			// exportExcelWorkSheet NOT FOUND
			LOGError( "EXCELEExportRow Could NOT export row because exportExcelWorkSheet NOT opened!" );
			return;
		}

		try {

			if ( exportColumns.length > 0 )
			{

				// Iterate over all columns specified in EXCELEExportInitialize()
				for ( let curCol = 1 ; curCol <= exportColumns.length ; curCol++ ) {
					// Get the column name
					let currentColumn = exportColumns[ curCol - 1 ];

					// Get the corresponding field value from valueMap
					let fieldValue = valueMap.get( currentColumn );

					// If the fieldValue is null/undefined, output an empty string
					if ( fieldValue == null ) {
						fieldValue = "";
					}
					
					exportExcelWorkSheet.Cells.Item( exportRow, curCol ).Value = __EXCELEToSafeEXCELString( fieldValue );
				}

				// Prepare the exportRow for the next export
				exportRow++;
				// Session.Output( "EXCELEExportRow exportRow = " + exportRow + "!!!" );
			}
		} catch (err) {
			LOGError( "EXCELEExportRow catched error " + err.message + "!" );
			return;
		}
	}
	else
	{
		LOGWarning( "EXCEL Export is not currently in progress. Call EXCELEExportInitialize() to start a EXCEL Export" );
	}

}

/**
 * Creates and returns an empty Value Map.
 */
function EXCELECreateEmptyValueMap() /* : Map */
{
	var valueMap = new Map();
	return valueMap;
}

/**
 * Returns an array of column names considered standard for EA elements. This array can be used
 * as the columns parameter when calling EXCELEExportInitialize()
 *
 * @param[in] includeGUID (boolean) Advises whether the elementGUID field should be included
 *
 * @return an array of column names 
 */
function EXCELEGetStandardElementColumns( includeGUID /* : boolean */ ) /* : Array */
{
	var columnArray = [];
	
	columnArray.push( "Abstract" );
	// ActionFlags - Not included (Comma Separated)	
	columnArray.push( "Alias" );
	// Attributes - Not included (Collection)
	// AttributesEx - Not included (Collection)
	columnArray.push( "Author" );
	// BaseClasses - Not included (Collection)
	// ClassifierID - Not included (Relational)
	columnArray.push( "ClassifierName" );
	// ClassifierType - Not included (Read Only)
	columnArray.push(  "Complexity" );
	// CompositeDiagram - Not included (Relational)
	// Connectors - Not included (Collection)
	// Constraints - Not included (Collection)
	// ConstraintsEx - Not included (Collection)
	columnArray.push( "Created" );
	// CustomProperties - Not included (Collection)
	// Diagrams - Not included (Collection)
	columnArray.push( "Difficulty" );
	// Efforts - Not included (Collection)
	
	if ( includeGUID ) {
		columnArray.push( "ElementGUID" );
	}
	
	// Elements - Not included (Collection)
	// EmbeddedElements - Not included (Collection)
	// EventFlags - Not included (Comma Separated)
	// ExtensionPoints - Not included (Comma Separated)
	// Files - Not included (Collection)
	columnArray.push( "GenFile" );
	// Genlinks - Not included (Relational)
	columnArray.push( "GenType" );
	columnArray.push( "Header1" );
	columnArray.push( "Header2" );
	columnArray.push( "IsActive" );
	columnArray.push( "IsLeaf" );
	columnArray.push( "IsNew" );
	columnArray.push( "IsSpec" );
	// Issues - Not included (Collection)
	columnArray.push( "Locked" );
	// MetaType - Not included (Read Only)
	// Methods - Not included (Collection)
	// Metrics - Not included (Collection)
	// MiscData - Not included (Read Only)
	// Modified - Not included (Overwritten)
	columnArray.push( "Multiplicity" );
	columnArray.push( "Name" );
	columnArray.push( "Notes" );
	// ObjectType - Not included (Read Only)
	// PackageID - Not included (Relational)
	// ParentID - Not included (Relational)
	// Partitions - Not included (Collection)
	columnArray.push( "Persistence" );
	columnArray.push( "Phase" );
	columnArray.push( "Priority" );
	// Properties - Not included (Collection)
	// PropertyType - Not included (Relational)
	// Realizes - Not included (Collection)
	// Requirements - Not included (Collection)
	// RequirementsEx - Not included (Collection)
	// Resources - Not included (Collection)
	// Risks - Not included (Collection)
	columnArray.push( "RunState" );
	// Scenarios - Not included (Collection)
	// State Transitions - Not included (Collection)
	columnArray.push( "Status" );
	columnArray.push( "Stereotype" );
	// StereotypeEx - Not included (Comma Separated)	
	columnArray.push( "Subtype" );
	columnArray.push( "Tablespace" );
	columnArray.push( "Tag" );
	// TaggedValues - Not included (Collection)
	// TaggedValuesEx - Not included (Collection)
	// Tests - Not included (Collection)
	columnArray.push( "TreePos" );
	columnArray.push( "Type" );
	columnArray.push( "Version" );
	columnArray.push( "Visibility" );
	
	return columnArray;
}

/**
 * Creates a Value Map of standard property names/values for the specified element. This Value Map 
 * can be used as the valueMap parameter when calling the ExportRow() function.
 *
 * @param[in] element (EA.Element) The element to compile the Value Map for
 *
 * @return A Value Map populated with the provided element's values.
 */
function EXCELEGetStandardElementFieldValues( element /* : EA.Element */ ) /* : Map */
{

	let valueMap = EXCELECreateEmptyValueMap();

	try {

		var theElement as EA.Element;
		theElement = element;

		valueMap.set( "Abstract", theElement.Abstract );
		// ActionFlags - Not included (Comma Separated)
		valueMap.set( "Alias", theElement.Alias );
		// Attributes - Not included (Collection)
		// AttributesEx - Not included (Collection)
		valueMap.set( "Author", theElement.Author );
		// BaseClasses - Not included (Collection)
		// ClassifierID - Not included (Relational)
		valueMap.set( "ClassifierName", theElement.ClassifierName );
		// ClassifierType - Not included (Read Only)
		valueMap.set( "Complexity", theElement.Complexity);
		// CompositeDiagram - Not included (Relational)
		// Connectors - Not included (Collection)
		// Constraints - Not included (Collection)
		// ConstraintsEx - Not included (Collection)
		valueMap.set( "Created", theElement.Created );
		// CustomProperties - Not included (Collection)
		// Diagrams - Not included (Collection)
		valueMap.set( "Difficulty", theElement.Difficulty );
		// Efforts - Not included (Collection)
		valueMap.set( "ElementGUID", theElement.ElementGUID );
		// Elements - Not included (Collection)
		// EmbeddedElements - Not included (Read Only)
		// Event Flags - Not included (Comma Separated)	
		// ExtensionPoints - Not included (Comma Separated)
		// Files - Not included (Collection)
		valueMap.set( "GenFile", theElement.Genfile );
		// Genlinks - Not included (Relational)
		valueMap.set( "GenType", theElement.Gentype );
		valueMap.set( "Header1", theElement.Header1 );
		valueMap.set( "Header2", theElement.Header2 );
		valueMap.set( "IsActive", theElement.IsActive );
		valueMap.set( "IsLeaf", theElement.IsLeaf );
		valueMap.set( "IsNew", theElement.IsNew );
		valueMap.set( "IsSpec", theElement.IsSpec );
		// Issues - Not included (Collection)
		valueMap.set( "Locked", theElement.Locked );
		// MetaType - Not included (Read Only)
		// Methods - Not included (Collection)
		// Metrics - Not included (Collection)
		// MiscData - Not included (Read Only)
		// Modified - Not included (Overwritten)
		valueMap.set( "Multiplicity", theElement.Multiplicity );
		valueMap.set( "Name", theElement.Name );
		valueMap.set( "Notes", theElement.Notes );
		// ObjectType - Not included (Read Only)
		// PackageID - Not included (Relational)
		// ParentID - Not included (Relational)
		// Partitions - Not included (Collection)
		valueMap.set( "Persistence", theElement.Persistence );
		valueMap.set( "Phase", theElement.Phase );
		valueMap.set( "Priority", theElement.Priority );
		// Properties - Not included (Collection)
		// PropertyType - Not included (Relational)
		// Realizes - Not included (Collection)
		// Requirements - Not included (Collection)
		// RequirementsEx - Not included (Collection)
		// Resources - Not included (Collection)
		// Risks - Not included (Collection)
		valueMap.set( "RunState", theElement.RunState );
		// Scenarios - Not included (Collection)
		// State Transitions - Not included (Collection)
		valueMap.set( "Status", theElement.Status );
		valueMap.set( "Stereotype", theElement.Stereotype );
		// StereotypeEx - Not included (Comma Separated)	
		valueMap.set( "Subtype", theElement.Subtype );
		valueMap.set( "Tablespace", theElement.Tablespace );
		valueMap.set( "Tag", theElement.Tag );
		// TaggedValues - Not included (Collection)
		// TaggedValuesEx - Not included (Collection)
		// Tests - Not included (Collection)
		valueMap.set( "TreePos", theElement.TreePos );
		valueMap.set( "Type", theElement.Type );
		valueMap.set( "Version", theElement.Version );
		valueMap.set( "Visibility", theElement.Visibility );

	} catch (err) {
		LOGError( "EXCELEGetStandardElementFieldValues catched error " + err.message + "!" );
		valueMap = null;
	}

	return valueMap;
}

/**
 * @private
 * Returns a copy of the string that is safe for inclusion in a EXCEL file.
 *
 * @param[in] originalString (String) The string to convert
 *
 * @return a copy of the string modified for inclusion in a EXCEL file
 */
function __EXCELEToSafeEXCELString( originalString /* : String */ ) /* : String */
{
	var returnString = new String(originalString);
	
	// Strip out delimiters
	var delimiterRegExp = new RegExp( EXCEL_DELIMITER, "gm" );
	returnString = returnString.replace( delimiterRegExp, "" );
	
	// Strip out newline chars
	var newlineRegExp = new RegExp( "\r\n?", "gm" );
	returnString = returnString.replace( newlineRegExp, " " );
		
	return returnString;
}
