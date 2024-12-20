//[group=BaatDiagramScripts]
!INC Local Scripts.EAConstants-JavaScript
!INC EAScriptLib.JavaScript-Dialog
!INC EAScriptLib.JavaScript-CSV
!INC BaatScriptLib.BaatScript-Logging
!INC BaatScriptLib.JavaScript-Connector
!INC BaatScriptLib.JavaScript-ImEx
!INC BaatScriptLib.ImportDataBricksCommon

/*
 * Script Name:	ImportDataBricksConnectors
 * Author:		J. de Baat
 * Purpose:		Import the Connector information from DataBricks CSV file into the repository using the BaatScriptLib scripts
 * Date:		20-12-2024
 * 
 */

const CSVFilterString        = "CSV Files (*.csv;*.txt)|*.csv;*.txt|All Files (*.*)|*.*||";

// The level to log at
var   BLOGLEVEL = BLOGLEVEL_WARNING;		// Choose from: BLOGLEVEL_ERROR, BLOGLEVEL_WARNING, BLOGLEVEL_INFO, BLOGLEVEL_DEBUG, BLOGLEVEL_TRACE

let   intImportRowNr = 0;
let   intShowDelta   = 10;
let   intShowRowNr   = 0;

/*
 * Handle the CSVImport for importing Connectors
 */
function IDBImportConnectors( )
{

	try {

		// Clear the cache before importing starts
		IDBElementCacheMap  = null;					// : Map   cache for IDBElements
		IDBElementCacheList = null;					// : Array cache for IDBElements

		// Get the CSV fileName for this Import session, true for readonly.
		let curCSVFileName = DLGOpenFile( CSVFilterString, 1 );
		if ( ( curCSVFileName == null ) || ( curCSVFileName == "" ) ) {
			BLOGError( "DLGOpenFile could NOT Get curCSVFileName!" );
			return " IDBImportConnectors DLGOpenFile could NOT Get curCSVFileName!!!";
		}
		strImExSource = ImExSourceCSV;

		// Start the CSVImport for this session
		Session.Output( _BLOGGetDisplayDate() + " IDBImportConnectors started CSV Import for curCSVFileName " + curCSVFileName + "!!!" );
		CSVIImportFile( curCSVFileName, true );

		// Clear the cache after importing finished
		IDBElementCacheMap  = null;					// : Map   cache for IDBElements
		IDBElementCacheList = null;					// : Array cache for IDBElements

		// Return "" as successful result
		return "";

	} catch (catch_err) {
		BLOGError( " catched error " + catch_err.message + "!" );
		return "IDBImportConnectors catched error " + catch_err.message + "!";
	}

}

/*
 * Process the Connectors found in the CSV Import file
 */
function OnRowImported( theRow )
{

	var curConnector as EA.Connector;

	let importRow = 0;

	try {

		importRow = importCurrentRow;
		intImportRowNr++;
		strMessage = "Processing importCurrentRow(" + importRow + ").length = " + importCurrentRow.length + ", [0] = " + importCurrentRow[0] + "!!!";
		BLOGTrace( strMessage );

		// Find and process the Connectors available in this importRow
		curConnector = IDBFindConnector( ImExColumnConnectorGUID, IDBTaggedValueIDBKey );
		if ( curConnector != null ) {
			// Process Connector found
			strMessage = "Found curConnector.Name = " + curConnector.Name + ", ConnectorID = " + curConnector.ConnectorID + ", ClientID = " + curConnector.ClientID + ", SupplierID = " + curConnector.SupplierID + "!!!";
			BLOGTrace( strMessage );
		} else {
			// Connector NOT found so create new Connector
			BLOGTrace("Needs to create new curConnector!!!" );
			curConnector = IDBCreateConnector();
		}

		// Process the ConnectorProperties
		if ( curConnector != null ) {
			BLOGTrace("Created curConnector.Name = " + curConnector.Name + ", ConnectorID = " + curConnector.ConnectorID + ", ClientID = " + curConnector.ClientID + ", SupplierID = " + curConnector.SupplierID + "!!!" );

			// Process the updates for curConnector found
			let curResult = IDBUpdateConnectorProperties( curConnector );
			strMessage = "Updated[ " + importRow + " ] Properties of ConnectorID " + curConnector.ConnectorID + ", curResult= " + curResult + "!!!";
			BLOGTrace( strMessage );

			// Only show progress for each intShowDelta Elements
			if ( intImportRowNr >= intShowRowNr ) {
				strMessage = _BLOGGetDisplayDate() + " Processed row[ " + intImportRowNr + " ] for ConnectorID " + curConnector.ConnectorID + ", curResult = "  + curResult + "!";
				Session.Output( strMessage );
				intShowRowNr += intShowDelta;
			}
		} else {
			strMessage = _BLOGGetDisplayDate() + "Processed row[ " + intImportRowNr + " ] for null curConnector!";
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
 * Create theConnector with information as provided
 */
function IDBCreateConnector()
{

	var curConnector       as EA.Connector;
	var curElementClient   as EA.Element;
	var curElementSupplier as EA.Element;

	curConnector            = null;

	try {

		// Find and process the source element defined by CSO_CLASSGUID or IDBTaggedValueIDBSource
		let curElementClient = IDBFindElement( ImExColumnCSOClassGUID, IDBTaggedValueIDBSource );
		if ( curElementClient == null ) {
			BLOGWarning("Could NOT find curElementClient so skip creation!!!");
			return null;
		}

		// Find and process the source element defined by CTO_CLASSGUID or IDBTaggedValueIDBTarget
		let curElementSupplier = IDBFindElement( ImExColumnCTOClassGUID, IDBTaggedValueIDBTarget );
		if ( curElementSupplier == null ) {
			BLOGWarning("Could NOT find curElementSupplier so skip creation!!!");
			return null;
		}

		// Create newConnector
		let curConnectorType = IDBIGetColumnValueByName( ImExColumnConnectorType );
		if ( curConnectorType == null ) {
			curConnectorType = ConnectorStereotypeDefault;
			BLOGTrace("Could NOT find curConnectorType so set to default: " + curConnectorType + " !!!");
		}
		BLOGTrace("Create new curConnector with Connector_Type " + curConnectorType + " for curElementClient " + curElementClient.Name + " and curElementSupplier = " + curElementSupplier.Name + "!!!" );
		curConnector = CONSetElementConnector( curElementClient, curElementSupplier, curConnectorType, true );

		// IDBCreateConnector created new curConnector so update it using the other values found
		if ( curConnector != null ) {

			// Set the curConnector Direction as Source -> Destination
			curConnector.ClientEnd.IsNavigable = false;
			curConnector.ClientEnd.Update();
			curConnector.SupplierEnd.IsNavigable = true;
			curConnector.SupplierEnd.Update();
			curConnector.Update();

			// Process the Properties for curConnector found
			curConnector.RouteStyle = ConnectorRouteStyleDefault;
			curConnector.Stereotype = ConnectorStereotypeDefault;
			curConnector.ClientEnd.Aggregation = ConnectorArchiMateAssociation;

			curConnector.Update();

		} else {

			strMessage = "Could NOT create curConnector for curElementClient " + curElementClient.Name + " and curElementSupplier = " + curElementSupplier.Name + " so NOT updated!!!";
			BLOGWarning( strMessage );
			return null;
		}

	} catch (catch_err) {
		BLOGError( " catched error " + catch_err.message + "!" );
	}

	return curConnector;
}

/*
 * Update theConnector with information as provided in the fields
 */
function IDBUpdateConnectorProperties( theConnector )
{

	// Cast theConnector to EA.Connector so we get intellisense
	var curConnector     as EA.Connector;
	curConnector          = theConnector;

	try {

		// Process theConnector
		if ( curConnector != null ) {

			BLOGTrace("( " + curConnector.ConnectorGUID + " ) found curConnector.ConnectorID = " + curConnector.ConnectorID + "!!!" );

			// Process StandardConnectorTaggedValues
			IDBISetConnectorTaggedValues( curConnector );
			BLOGTrace("( " + curConnector.ConnectorGUID + " ) updated curConnector.ConnectorID to " + curConnector.ConnectorID + ", #curConnectorTags = " + curConnector.TaggedValues.Count + "!!!" );

			// Add TaggedValues to the curConnector
			TVSetElementTaggedValue( curConnector, IDBTaggedValueIDBProcessed,   _BLOGGetDisplayDate(),    true );
			TVSetElementTaggedValue( curConnector, IDBTaggedValueIDBConnectorID, curConnector.ConnectorID, true );
			TVSetElementTaggedValue( curConnector, IDBTaggedValueIDBSourceID,    curConnector.ClientID,    true );
			TVSetElementTaggedValue( curConnector, IDBTaggedValueIDBTargetID,    curConnector.SupplierID,  true );
			BLOGTrace("( " + curConnector.ConnectorGUID + " ) added TaggedValues " + IDBTaggedValueIDBProcessed + " ( " + _BLOGGetDisplayDate() + " ), " + IDBTaggedValueIDBSource + " ( " + curConnector.ClientID + " )!!!" );

			// Commit the changes to the repository
			objGlobalIDBPackage.Update();
			objGlobalIDBPackage.Connectors.Refresh();

		} else {

			return "Could NOT find curConnector so NOT updated!!!";
		}

	} catch (catch_err) {
		BLOGError( " catched error " + catch_err.message + "!" );
		return "IDBUpdateConnectorProperties catched error " + catch_err.message + "!";
	}

	return "";
}

/*
 * Diagram Script main function
 */
function ImportDataBricksConnectors()
{
	// Show the script output window
	Repository.EnsureOutputVisible( "Script" );

	Session.Output( "======================================= Started ImportDataBricksConnectors at " + _BLOGGetDisplayDate() + "!" );

	try
	{

		// Get and check the global variables
		const validDiagram = IDBGetAndCheckDiagram();
		if ( validDiagram ) {
			BLOGInfo( "Diagram is VALID so proceed processing!" );
			strImExStart       = strImExStartDiagram;
			objGlobalEADiagram = objGlobalIDBDiagram;

			let curResult = IDBImportConnectors();
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

	Session.Output( "======================================= Finished ImportDataBricksConnectors at " + _BLOGGetDisplayDate() + "!" );
}

ImportDataBricksConnectors();
