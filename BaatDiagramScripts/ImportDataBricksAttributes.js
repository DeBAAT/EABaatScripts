//[group=BaatDiagramScripts]
!INC Local Scripts.EAConstants-JavaScript
!INC EAScriptLib.JavaScript-Dialog
!INC EAScriptLib.JavaScript-CSV
!INC BaatScriptLib.BaatScript-Logging
!INC BaatScriptLib.JavaScript-Connector
!INC BaatScriptLib.JavaScript-ImEx
!INC BaatScriptLib.ImportDataBricksCommon

/*
 * Script Name:	ImportDataBricksAttributes
 * Author:		J. de Baat
 * Purpose:		Import the Attribute information from DataBricks CSV file into the repository using the BaatScriptLib scripts
 * Date:		20-12-2024
 * 
 */

const CSVFilterString        = "CSV Files (*.csv;*.txt)|*.csv;*.txt|All Files (*.*)|*.*||";

// The level to log at
var   BLOGLEVEL = BLOGLEVEL_WARNING;		// Choose from: BLOGLEVEL_ERROR, BLOGLEVEL_WARNING, BLOGLEVEL_INFO, BLOGLEVEL_DEBUG, BLOGLEVEL_TRACE
let   intImportRowNr = 0;
let   intShowDelta   = 100;
let   intShowRowNr   = 0;

/*
 * Handle the CSVImport for importing Attributes
 */
function IDBImportAttributes( )
{

	try {

		// Clear the cache before importing starts
		IDBElementCacheMap  = null;					// : Map   cache for IDBElements
		IDBElementCacheList = null;					// : Array cache for IDBElements

		// Get the CSV fileName for this Import session, true for readonly.
		let curCSVFileName = DLGOpenFile( CSVFilterString, 1 );
		if ( ( curCSVFileName == null ) || ( curCSVFileName == "" ) ) {
			BLOGError( "DLGOpenFile could NOT Get curCSVFileName!" );
			return " IDBImportAttributes DLGOpenFile could NOT Get curCSVFileName!!!";
		}
		strImExSource = ImExSourceCSV;

		// Start the CSVImport for this session
		Session.Output( _BLOGGetDisplayDate() + " IDBImportAttributes started CSV Import for curCSVFileName " + curCSVFileName + "!!!" );
		CSVIImportFile( curCSVFileName, true );

		// Clear the cache after importing finished
		IDBElementCacheMap  = null;					// : Map   cache for IDBElements
		IDBElementCacheList = null;					// : Array cache for IDBElements

		// Return "" as successful result
		return "";

	} catch (catch_err) {
		BLOGError( " catched error " + catch_err.message + "!" );
		return "IDBImportAttributes catched error " + catch_err.message + "!";
	}

}

/*
 * Process the Attributes found in the CSV Import file
 */
function OnRowImported( theRow )
{

	var curElement as EA.Element;

	let importRow = 0;

	try {

		importRow = importCurrentRow;
		intImportRowNr++;
		strMessage = " importCurrentRow(" + importRow + ").length = " + importCurrentRow.length + ", Action = " + importCurrentRow[0] + "!!!";
		BLOGTrace( strMessage );

		// Find and process the elements available in this importRow
		curElement = IDBFindElement( ImExColumnClassGUID, IDBTaggedValueIDBKey );
		if ( curElement != null ) {
			// Process element found
			strMessage = "Processing curElement.Name = " + curElement.Name + ", ElementID = " + curElement.ElementID + ", ParentID = " + curElement.ParentID + ", PackageID = " + curElement.PackageID + "!!!";
			BLOGTrace( strMessage );

			// Create the Attribute for curElement found
			let curResult = IDBCreateElementAttribute( curElement );
			if ( intImportRowNr >= intShowRowNr ) {
				strMessage = _BLOGGetDisplayDate() + " Processed row[ " + intImportRowNr + " ] for " + curElement.Name + ", curResult = "  + curResult + "!";
				Session.Output( strMessage );
				intShowRowNr += intShowDelta;
			}
		} else {
			strMessage = _BLOGGetDisplayDate() + "Processed row[ " + intImportRowNr + " ] for null curElement!";
			Session.Output( strMessage );
		}

	} catch (catch_err) {
		BLOGError( " catched error " + catch_err.message + "!" );
	}

	// Free up memory
	curElement        = null;
	currentLine       = null;
	currentLineTokens = null;
}

/*
 * Create the Attribute for curElement found with information as provided
 */
function IDBCreateElementAttribute( theElement )
{

	// Cast theElement to EA.Element so we get intellisense
	var curElement     as EA.Element;
	var curAttributes  as EA.Collection;
	var curAttribute   as EA.Attribute;

	// Get and check the ImExColumnAttributeName parameter
	let newAttributeName = IDBIGetColumnValueByName( ImExColumnAttributeName );
	if ( ( newAttributeName == null ) || ( newAttributeName == "" ) ) {
		// Return null because not found
		return "IDBProcessElementAttribute: newAttributeName not available so Attribute NOT created!!!";
	}

	// Check all Attributes in theTypePackage whether the requested theAttributeName already exists
	curElement    = theElement;
	curAttributes = curElement.Attributes;
	curAttribute  = IDBGetCollectionObjectByName( curAttributes, newAttributeName );

	// Create new Attribute when it does not exist yet
	if ( curAttribute == null ) {

		// Get and check the ImExCSVColumnAttributeType parameter
		let newAttributeType = IDBIGetColumnValueByName( ImExColumnAttributeType );
		if ( ( newAttributeType == null ) || ( newAttributeType == "" ) ) {
			// Use default value because not found
			newAttributeType = "string";
			BLOGTrace("Use default value string for newAttributeType because not found!!!" );
		}

		// Create the new curAttribute
		curAttribute = curAttributes.AddNew( newAttributeName, newAttributeType );
		curAttribute.Update();
		curAttributes.Refresh();
	}

	strMessage = "Started curAttribute[ " + curAttribute.Name + " ] !";
	BLOGTrace( strMessage );

	// Update the parameters for curAttribute
	// curAttribute = IDBUpdateAttributeProperties( curAttribute );
	// curAttributes.Refresh();

	let curResult = "IDBProcessElementAttribute: Processed " + curAttribute.Name + " for Type " + curAttribute.Type + " !!!";

	// Free up memory
	curElement    = null;
	curAttribute  = null;
	curAttributes = null;

	return curResult;

}

/*
 * Update theAttribute with information as provided
 */
function IDBUpdateAttributeProperties( theAttribute )
{

	// Cast theAttribute to EA.Attribute so we get intellisense
	var curAttribute as EA.Attribute;

	curAttribute = theAttribute;

	// Check information found
	let newAttributeNotes = IDBIGetColumnValueByName( ImExColumnNotes );
	if ( ( newAttributeNotes == null ) || ( newAttributeNotes == "" ) ) {
		// Define newAttributeNotes because not found
		newAttributeNotes = curAttribute.Name + " created by the ImportDataBricksAttributes script for Type " + curAttribute.Type + " on " + _LOGGetDisplayDate() + ".";
	}

	// IsCollection indicates if the current feature is a collection or not. If the attribute represents a database column this, when set, represents a ForeignKey.
	let newAttributeIsCollection = IDBIGetColumnValueByName( ImExColumnIsCollection );
	if ( ( newAttributeIsCollection == null ) || ( newAttributeIsCollection == "" ) ) {
		// Get newAttributeIsCollection from ImExCSVColumnForeignKey because ImExCSVColumnIsCollection not found
		newAttributeIsCollection = IDBIGetColumnValueByName( ImExColumnForeignKey );
	}
	if ( ( newAttributeIsCollection != null ) && ( newAttributeIsCollection != "" ) ) {
		curAttribute.IsCollection = newAttributeIsCollection;
	}

	// IsOrdered indicates if a collection is ordered or not. If the attribute represents a database column this, when set, represents a PrimaryKey.
	let newAttributeIsOrdered = IDBIGetColumnValueByName( ImExColumnIsOrdered );
	if ( ( newAttributeIsOrdered == null ) || ( newAttributeIsOrdered == "" ) ) {
		// Get newAttributeIsOrdered from ImExCSVColumnPrimaryKey because ImExCSVColumnIsOrdered not found
		newAttributeIsOrdered = IDBIGetColumnValueByName( ImExColumnPrimaryKey );
	}
	if ( ( newAttributeIsOrdered != null ) && ( newAttributeIsOrdered != "" ) ) {
		curAttribute.IsOrdered = newAttributeIsOrdered;
	}

	// IsStatic indicates if the current attribute is a static feature or not. If the attribute represents a database column this, when set, represents the 'Unique' option.
	let newAttributeIsStatic  = IDBIGetColumnValueByName( ImExColumnIsStatic );
	if ( ( newAttributeIsStatic == null ) || ( newAttributeIsStatic == "" ) ) {
		// Get newAttributeIsStatic from ImExCSVColumnUnique because ImExCSVColumnIsStatic not found
		newAttributeIsStatic = IDBIGetColumnValueByName( ImExColumnUnique );
	}
	if ( ( newAttributeIsStatic != null ) && ( newAttributeIsStatic != "" ) ) {
		curAttribute.IsStatic = newAttributeIsStatic;
	}

	// Length indicates the attribute length, where applicable.
	let newAttributeLength = IDBIGetColumnValueByName( ImExColumnLength );
	if ( ( newAttributeLength != null ) && ( newAttributeLength != "" ) ) {
		curAttribute.Length = newAttributeLength;
	}

	strMessage = "Started newAttribute[ " + curAttribute.Name + " ] for IsOrdered " + curAttribute.IsOrdered + ", newAttributeIsOrdered " + newAttributeIsOrdered + " !";
	BLOGTrace( strMessage );

	// Update the parameters for curAttribute
	curAttribute.Notes = newAttributeNotes;
	curAttribute.Update();

	strMessage = "Updated newAttribute[ " + curAttribute.Name + " ] for Type " + curAttribute.Type + ", IsCollection " + curAttribute.IsCollection + " !";
	BLOGTrace( strMessage );

	return curAttribute;
}

/*
 * Diagram Script main function
 */
function ImportDataBricksAttributes()
{
	// Show the script output window
	Repository.EnsureOutputVisible( "Script" );

	Session.Output( "======================================= Started ImportDataBricksAttributes at " + _BLOGGetDisplayDate() + "!" );

	try
	{

		// Get and check the global variables
		const validDiagram = IDBGetAndCheckDiagram();
		if ( validDiagram ) {
			BLOGInfo( "Diagram is VALID so proceed processing!" );
			strImExStart       = strImExStartDiagram;
			objGlobalEADiagram = objGlobalIDBDiagram;

			let curResult = IDBImportAttributes();
			if ( curResult.length > 0 ) {
				BLOGError( curResult );
			} else {
				BLOGInfo( "Finished processing!" );
			}

			// Refresh the objGlobalIDBDiagram using ReloadDiagram
			objGlobalIDBDiagram.Update();
			Repository.ReloadDiagram( objGlobalIDBDiagram.DiagramID );

		} else {
			BLOGError( "Diagram is NOT VALID!" );
		}
	}
	catch(catch_err)
	{
		BLOGError(" found Error: " + catch_err + "!!!" );
	}

	Session.Output( "======================================= Finished ImportDataBricksAttributes at " + _BLOGGetDisplayDate() + "!" );
}

ImportDataBricksAttributes();
