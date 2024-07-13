//[group=BaatPackageScripts]
!INC Local Scripts.EAConstants-JavaScript
!INC EAScriptLib.JavaScript-Logging
!INC EAScriptLib.JavaScript-Dialog
!INC EAScriptLib.JavaScript-XML
!INC BaatScriptLib.JavaScript-Connector
!INC BaatScriptLib.JavaScript-EXCEL
!INC BaatScriptLib.JavaScript-ImEx

/*
 * This code has been included from the default Project Browser template.
 * If you wish to modify this template, it is located in the Config\Script Templates
 * directory of your EA install path.   
 * 
 * Script Name:	ImExImportConnectorsJAVA
 * Author:		J de Baat
 * Purpose:		Import the information from Connectors into the selected Package
 * Date:		13-07-2024
 * 
 */

/*
 * Handle the ExcelImport for importing Elements
 */
function ImExImportConnectors( )
{

	// Start the ExcelImport for this Import session
	// Session.Output("ImExImportElements started Excel.Application !" );
	let strExcelImportResult = IMEXIHandleExcelImport( strDefaultConnectorsSheetName );

	// Return the result found
	return strExcelImportResult;

}

/*
 * Process the Imported Connectors found
 */
function OnExcelRowImported( theRow )
{

	let actionResult = "";
	let importAction = "";
	let importRow    = 0;

	if ( excelImportCurrentRow[0] != null ) {
		importAction = excelImportCurrentRow[0].toLowerCase();
	}
	importRow = theRow;
	// Session.Output("OnExcelRowImported excelImportCurrentRow(" + importRow + ").length = " + excelImportCurrentRow.length + ", Action = " + excelImportCurrentRow[0] + "!!!" );

	// Process importAction requested
	switch( importAction )
	{
		case "create":
		{
			 actionResult = createConnector();
			 break;
		}
		case "update":
		{
			 actionResult = updateConnector();
			 break;
		}
		case "delete":
		{
			 actionResult = deleteConnector();
			 break;
		}
		case "":
		{
			 actionResult = "OnExcelRowImported found empty importAction!";
			 break;
		}
		default:
		{
			// Error message
			LOGError( "OnExcelRowImported CANNOT process importAction " + importAction + "!"  );
		}
	}

}

/*
 * Update theConnector with information as provided
 */
function createConnector()
{

	var curConnector       as EA.Connector;
	var curElementClient   as EA.Element;
	var curElementSupplier as EA.Element;

	// Find and process the element found by curConnectorGUID
	let curConnectorGUID = EXCELIGetColumnValueByName("CONNECTORGUID");
	curConnector         = GetConnectorByGuid( curConnectorGUID );
	if ( curConnector != null ) {
		return "createConnector( " + curConnectorGUID + " ) found curConnector.ConnectorID = " + curConnector.ConnectorID + " so skip creation to prevent duplicates!!!" ;
	} else {

		// Find and process the source element defined by Start_Object_ID
		let curElementClientID = EXCELIGetColumnValueByName("Start_Object_ID");
		curElementClient       = GetElementByID( curElementClientID );
		if ( curElementClient == null ) {
			// Session.Output("createConnector( " + curConnectorGUID + " ) could NOT find curElementClient with curElementClientID = " + curElementClientID + " so skip creation!!!" );
			return "createConnector( " + curConnectorGUID + " ) could NOT find curElementClient with curElementClientID = " + curElementClientID + " so skip creation!!!";
		}

		// Find and process the target element defined by End_Object_ID
		let curElementSupplierID = EXCELIGetColumnValueByName("End_Object_ID");
		curElementSupplier       = GetElementByID( curElementSupplierID );
		if ( curElementSupplier == null ) {
			// Session.Output("createConnector( " + curConnectorGUID + " ) could NOT find curElementSupplier with curElementSupplierID = " + curElementSupplierID + " so skip creation!!!" );
			return "createConnector( " + curConnectorGUID + " ) could NOT find curElementSupplier with curElementSupplierID = " + curElementSupplierID + " so skip creation!!!";
		}

		// createConnector could NOT find curConnector so create new one
		let curConnectorType = EXCELIGetColumnValueByName("Connector_Type");
		// Session.Output("createConnector( " + curConnectorGUID + " ) could NOT find curConnector so create new one with Connector_Type " + curConnectorType + "!!!" );
		curConnector = CONSetElementConnector( curElementClient, curElementSupplier, curConnectorType, false );

		// createConnector created new curConnector so update it using the other values found
		if ( curConnector != null ) {

			// Process the updates for curConnector found
			let curResult = updateConnectorProperties( curConnector );
			// Session.Output("createConnector( " + curConnectorGUID + " ) updated curConnector.Name to " + curConnector.Name + ", curResult= " + curResult + "!!!" );
			return curResult;

		} else {

			// Session.Output("createConnector( " + curConnectorGUID + " ) could NOT create curConnector so NOT updated!!!" );
			return "createConnector( " + curConnectorGUID + " ) could NOT create curConnector so NOT updated!!!";
		}
	}

	return "";
}

/*
 * Update theConnector with information as provided
 */
function updateConnector()
{

	var curConnector       as EA.Connector;
	var curElementClient   as EA.Element;
	var curElementSupplier as EA.Element;
	var newElement         as EA.Element;

	// Find and process the element found by curConnectorGUID
	let curConnectorGUID = EXCELIGetColumnValueByName("CONNECTORGUID");
	// Session.Output("updateConnector looking for curConnectorGUID: " + curConnectorGUID + "!!!" );
	if ( curConnectorGUID == null ) {
		// Session.Output("updateConnector could NOT find curConnectorGUID: " + curConnectorGUID + " so skip update!!!" );
		return "updateConnector could NOT find curConnectorGUID: " + curConnectorGUID + " so skip update!!!";
	}

	curConnector         = GetConnectorByGuid( curConnectorGUID );
	if ( curConnector != null ) {

		// Process the updates for curConnector found
		// Session.Output("updateConnector( " + curConnectorGUID + " ) found curConnector.Name = " + curConnector.Name + ", Visibility = " + curConnector.Visibility + "!!!" );
		let curResult = updateConnectorProperties( curConnector );
		// Session.Output("updateConnector( " + curConnectorGUID + " ) updated curConnector.Name to " + curConnector.Name + ", curResult= " + curResult + "!!!" );

		// Find and process the source element defined by Start_Object_ID
		let curElementClientID = EXCELIGetColumnValueByName("Start_Object_ID");
		curElementClient       = GetElementByID( curElementClientID );
		if ( curElementClient == null ) {
			// Session.Output("updateConnector( " + curConnectorGUID + " ) could NOT find curElementClient with curElementClientID = " + curElementClientID + " so skip update!!!" );
			return "updateConnector( " + curConnectorGUID + " ) could NOT find curElementClient with curElementClientID = " + curElementClientID + " so skip update!!!";
		}

		// Check whether to Change the ClientID when Start_Object_ID changed
		newElement = IMEXIGetNewConnectorStartOrEnd( curConnector, "Start_Object_ID", curElementClient );
		if ( newElement != null ) {
			curConnector.ClientID = newElement.ElementID;
			curConnector.Update();
			newElement.Update();
			curElementClient.Update();
		}

		// Find and process the target element defined by End_Object_ID
		let curElementSupplierID = EXCELIGetColumnValueByName("End_Object_ID");
		curElementSupplier       = GetElementByID( curElementSupplierID );
		if ( curElementSupplier == null ) {
			// Session.Output("updateConnector( " + curConnectorGUID + " ) could NOT find curElementSupplier with curElementSupplierID = " + curElementSupplierID + " so skip update!!!" );
			return "updateConnector( " + curConnectorGUID + " ) could NOT find curElementSupplier with curElementSupplierID = " + curElementSupplierID + " so skip update!!!";
		}

		// Check whether to Change the SupplierID when End_Object_ID changed
		newElement = IMEXIGetNewConnectorStartOrEnd( curConnector, "End_Object_ID", curElementSupplier );
		if ( newElement != null ) {
			curConnector.SupplierID = newElement.ElementID;
			curConnector.Update();
			newElement.Update();
			curElementSupplier.Update();
		}
		// Session.Output("updateConnectorProperties( " + curConnector.ConnectorGUID + " ) updated curConnector.Name to " + curConnector.Name + "!!!" );

	} else {

		// Session.Output("updateConnector( " + curConnectorGUID + " ) could NOT find curConnector so NOT updated!!!" );
		return "updateConnector( " + curConnectorGUID + " ) could NOT find curConnector so NOT updated!!!";
	}

	return "";
}

/*
 * Update theConnector with information as provided in the fields
 */
function updateConnectorProperties( theConnector )
{

	// Cast theConnector to EA.Connector so we get intellisense
	var curConnector as EA.Connector;

	curConnector      = theConnector;

	// Process theConnector
	if ( curConnector != null ) {

		// Process StandardConnectorFieldValues
		// Session.Output("updateConnectorProperties( " + curConnector.ConnectorGUID + " ) found curConnector.Name = " + curConnector.Name + "!!!" );
		IMEXISetStandardConnectorFieldValues( curConnector );
		// Session.Output("updateConnectorProperties( " + curConnector.ConnectorGUID + " ) updated curConnector.Name to " + curConnector.Name + "!!!" );

		// Commit the changes to the repository
		objGlobalEAPackage.Update();
		objGlobalEAPackage.Elements.Refresh();

	} else {

		return "updateConnectorProperties() could NOT find curConnector so NOT updated!!!" ;
	}

	return "";
}

/*
 * Delete theConnector with information as provided
 */
function deleteConnector()
{

	// Find and process the element found by curConnectorGUID
	let curConnectorGUID = EXCELIGetColumnValueByName("CONNECTORGUID");
	// Session.Output("deleteConnector looking for curConnectorGUID: " + curConnectorGUID + "!!!" );
	if ( curConnectorGUID == null ) {
		// Session.Output("deleteConnector could NOT find curConnectorGUID: " + curConnectorGUID + " so NOT deleted!!!" ;
		return "deleteConnector could NOT find curConnectorGUID: " + curConnectorGUID + " so NOT deleted!!!" ;
	}

	// Delete curConnectorGUID found
	// Session.Output("deleteConnector starting for curConnectorGUID = " + curConnectorGUID + "!!!" );
	let curResult = CONDeleteConnectorByGUID( curConnectorGUID );
	// Session.Output("deleteConnector starting for curConnectorGUID = " + curConnectorGUID + " resulted in: " + curResult + "!!!" );
	return curResult;

}
