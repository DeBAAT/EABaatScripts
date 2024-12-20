//[group=BaatDiagramScripts]
!INC Local Scripts.EAConstants-JavaScript
!INC EAScriptLib.JavaScript-Dialog
!INC EAScriptLib.JavaScript-CSV
!INC BaatScriptLib.BaatScript-Logging
!INC BaatScriptLib.JavaScript-Connector
!INC BaatScriptLib.JavaScript-ImEx
!INC BaatScriptLib.ImportDataBricksCommon

/*
 * This code has been included from the default Diagram Script template.
 * If you wish to modify this template, it is located in the Config\Script Templates
 * directory of your EA install path.
 *
 * Script Name:	ImportDataBricksElements
 * Author:		J. de Baat
 * Purpose:		Import the Element information from DataBricks CSV file into the repository using the BaatScriptLib scripts
 * Date:		20-12-2024
 * 
 */

const CSVFilterString        = "CSV Files (*.csv;*.txt)|*.csv;*.txt|All Files (*.*)|*.*||";

// The level to log at
var   BLOGLEVEL = BLOGLEVEL_WARNING;		// Choose from: BLOGLEVEL_ERROR, BLOGLEVEL_WARNING, BLOGLEVEL_INFO, BLOGLEVEL_DEBUG, BLOGLEVEL_TRACE
let   intImportRowNr = 0;
let   intShowDelta   = 10;					// Was 500 for 8000 attributes;
let   intShowRowNr   = 0;

/*
 * Handle the CSVImport for importing Elements
 */
function IDBImportElements( )
{

	try {

		// Clear the cache before importing starts
		IDBElementCacheMap  = null;					// : Map   cache for IDBElements
		IDBElementCacheList = null;					// : Array cache for IDBElements

		// Get the CSV fileName for this Import session, true for readonly.
		let curCSVFileName = DLGOpenFile( CSVFilterString, 1 );
		if ( ( curCSVFileName == null ) || ( curCSVFileName == "" ) ) {
			BLOGError( "DLGOpenFile could NOT Get curCSVFileName!" );
			return " IDBImportElements DLGOpenFile could NOT Get curCSVFileName!!!";
		}
		strImExSource = ImExSourceCSV;

		// Start the CSVImport for this session
		Session.Output( _BLOGGetDisplayDate() + " IDBImportElements started CSV Import for curCSVFileName " + curCSVFileName + "!!!" );
		CSVIImportFile( curCSVFileName, true );

		// Clear the cache after importing finished
		IDBElementCacheMap  = null;					// : Map   cache for IDBElements
		IDBElementCacheList = null;					// : Array cache for IDBElements

		// Return "" as successful result
		return "";

	} catch (catch_err) {
		BLOGError( " catched error " + catch_err.message + "!" );
		return "IDBImportElements catched error " + catch_err.message + "!";
	}

}

/*
 * Process the Elements found in the CSV Import file
 */
function OnRowImported( theRow )
{

	var curElement as EA.Element;

	let importRow = 0;

	try {

		importRow = importCurrentRow;
		intImportRowNr++;
		strMessage = "Processing importCurrentRow(" + importRow + ").length = " + importCurrentRow.length + ", [0] = " + importCurrentRow[0] + "!!!";
		BLOGTrace( strMessage );

		// Find and process the elements available in this importRow
		curElement = IDBFindElement( ImExColumnClassGUID, IDBTaggedValueIDBKey );
		if ( curElement != null ) {
			// Process element found
			strMessage = "Processing curElement.Name = " + curElement.Name + ", ElementID = " + curElement.ElementID + ", ParentID = " + curElement.ParentID + ", PackageID = " + curElement.PackageID + "!!!";
			BLOGTrace( strMessage );
		} else {
			// Element NOT found so create new element
			curElement = IDBCreateElement();
		}

		// Process the ElementProperties
		if ( curElement != null ) {

			// Process the Elements for curElement found
			let curResult = IDBUpdateElementProperties( curElement );

			// Only show progress for each intShowDelta Elements
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
 * Create theElement with information as provided
 */
function IDBCreateElement()
{

	var curElement      as EA.Element;
	var curStagePackage as EA.Package;

	curElement           = null;
	curStagePackage      = null;

	try {
		// createElement could NOT find curElement so create new one
		let curElementName = IDBIGetColumnValueByName( ImExColumnName      );
		let curElementType = IDBIGetColumnValueByName( ImExColumnClassType );
		if ( ( curElementName == null ) || ( curElementType == null ) ) {
			// Return null because not found
			BLOGWarning(" curElementName AND curElementType not available so Element NOT created!!!" );
			return null;
		}

		// Get and check theGlobalPathPackage
		let curStageName = IDBIGetColumnValueByName( IDBColumnIDBStageName );
		if ( curStageName == null ) {
			// Return null because not found
			BLOGWarning(" curStageName not available so Element NOT created!!!" );
			return null;
		}
		curStagePackage = IDBCheckOrAddSubPackage( objGlobalIDBPackage, curStageName );
		if ( curStagePackage == null ) {
			BLOGWarning(" curStagePackage( " + curStageName + " ) not available so set to objGlobalIDBPackage( " + objGlobalIDBPackage.Name + " ) !!!" );
			curStagePackage = objGlobalIDBPackage;
		}
		BLOGTrace("( " + curElementName + " ) creating new Element with Type " + curElementType + " in curStagePackage( " + curStageName + " )!!!" );
		curElement = curStagePackage.Elements.AddNew( curElementName, curElementType );

		// IDBCreateElement created new curElement so update it using the other values found
		if ( curElement != null ) {
			// Commit the changes to the repository
			curStagePackage.Update();
			curStagePackage.Elements.Refresh();

		} else {

			BLOGWarning("( " + curElementName + " ) could NOT create curElement in curStagePackage( " + curStagePackage.Name + " ) so NOT updated!!!");
			return null;
		}

	} catch (catch_err) {
		BLOGError( " catched error " + catch_err.message + "!" );
	}

	return curElement;
}

/*
 * Update theElement with information as provided
 */
function IDBUpdateElementProperties( theElement )
{

	// Cast theElement to EA.Element so we get intellisense
	var curElement as EA.Element;

	try {

		curElement = theElement;

		// Process theElement
		if ( curElement != null ) {

			// Check information found
			let newElementNotes = IDBIGetColumnValueByName( ImExColumnNotes );
			if ( ( newElementNotes == null ) || ( newElementNotes.trim() == "" ) ){
				// Define newElementNotes because not found
				newElementNotes = curElement.Name + " created by the ImportDataBricksElements script for Type " + curElement.Type + " on " + _LOGGetDisplayDate() + ".";
			}

			// Process StandardElementTaggedValues
			IDBISetElementTaggedValues( curElement );
			strMessage = "( " + curElement.ElementGUID + " ) updated curElement.Name " + curElement.Name + " with #curElementTags = " + curElement.TaggedValues.Count + "!!!";
			BLOGTrace( strMessage );

			// If TaggedValue in import, add it to the curElement
			TVSetElementTaggedValue( curElement, IDBTaggedValueIDBProcessed, _BLOGGetDisplayDate(), true );
			strMessage = "( " + curElement.ElementGUID + " ) added TaggedValue " + IDBTaggedValueIDBProcessed + " ( " + _BLOGGetDisplayDate() + ")!!!";
			BLOGTrace( strMessage );

			// Update the parameters for curElement
			curElement.Notes = newElementNotes;
			curElement.Update();

			strMessage = "Updated newElement[ " + curElement.Name + " ] for Type " + curElement.Type + ", curElement.Notes " + curElement.Notes + " !";
			BLOGTrace( strMessage );

		} else {

			return "IDBUpdateElementProperties() could NOT find curElement so NOT updated!!!";
		}

	} catch (catch_err) {
		BLOGError( " catched error " + catch_err.message + "!" );
		return " catched error " + catch_err.message + "!" ;
	}

	return "";
}

/*
 * Diagram Script main function
 */
function ImportDataBricksElements()
{
	// Show the script output window
	Repository.EnsureOutputVisible( "Script" );

	Session.Output( "======================================= Started ImportDataBricksElements at " + _BLOGGetDisplayDate() + "!" );

	try
	{

		// Get and check the global variables
		const validDiagram = IDBGetAndCheckDiagram();
		if ( validDiagram ) {
			BLOGInfo( "Diagram is VALID so proceed processing!" );
			strImExStart       = strImExStartDiagram;
			objGlobalEADiagram = objGlobalIDBDiagram;

			// Import the Elements
			let curResult = IDBImportElements();
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

	Session.Output( "======================================= Finished ImportDataBricksElements at " + _BLOGGetDisplayDate() + "!" );
}

ImportDataBricksElements();
