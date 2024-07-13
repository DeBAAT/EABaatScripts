//[group=BaatScriptLib]
!INC EAScriptLib.JavaScript-Logging

/**
 * @file JavaScript-Connector 
 * This script library contains helper functions for working with Connectors. Functions 
 * provided by this module are identified by the prefix CON.
 *
 * @author J. de Baat, based on JavaScript-TaggedValue by Sparx Systems
 * @date 2024-07-13
 */

/**
 * Retrieves the Connector object from the provided element whose data matches the specified parameters.
 * If the element does not exist, or does not contain a Connector with the specified data, null
 * is returned.
 * NOTE: A Connector is found if the ClientID AND SupplierID and Type are identical.
 *
 * @param[in] theElement (EA.Element) The element to retrieve the Connector value from
 * @param[in] theClientID (String) The ID of the Element registered as Source of the Connector
 * @param[in] theSupplierID (String) The ID of the Element registered as Target of the Connector
 * @param[in] theType (String) The type of the Connector to be found, if empty string, any type will match
 *
 * @return The object of the requested Connector found, null when not found
 */
function CONGetElementConnectorByData( theElement /* : EA.Element */, theClientID /* : ID */, theSupplierID /* : ID */, theType /* : String */ ) /* : EA.Connector */
{

	// Validate input parameters
	if ( ( theElement != null ) && ( theClientID > 0 ) ) {

		// Cast the input values to objects so we get intellisense
		var curElement           as EA.Element;
		var curElementConnectors as EA.Collection;
		var curConnector         as EA.Connector;

		curElement           = theElement;
		curElementConnectors = curElement.Connectors;

		// Check all element Connectors against data to find
		for ( var i = 0 ; i < curElementConnectors.Count ; i++ )
		{
			curConnector = curElementConnectors.GetAt( i );
			if ( ( curConnector.ClientID   == theClientID ) &&
				 ( curConnector.SupplierID == theSupplierID ) &&
				 ( ( "" === theType ) || ( curConnector.Type == theType ) ) ) 
			{
				// Connector found so clean up memory and return object
				curElementConnectors = null;
				return curConnector;
			}
		}

		// Clean up memory
		curElementConnectors = null;
	}

	return null;

}

/**
 * Sets the specified Connector on the provided element. If the provided element does not already
 * contain a Connector with the specified data, a new Connector is created.
 * If a Connector already exists with the specified data then the action is ignored.
 * NOTE: A Connector is found if the ClientID AND SupplierID and Type are identical.
 *
 * @param[in] theElementClient (EA.Element) The element to add the Connector to and registered as Source of the Connector
 * @param[in] theElementSupplier (EA.Element) The element to be registered as Target of the Connector
 * @param[in] theType (String) The type of the Connector to be found, if empty string, any type will match
 * @param[in] skipDuplicate (boolean) If set to true, check for existing connector to prevent duplicate
 *
 * @return The object of the Connector added or found, null in case of error
 */
function CONSetElementConnector( theElementClient /* : EA.Element */, theElementSupplier /* : EA.Element */, theType /* : String */, skipDuplicate /* : Boolean */ ) /* : EA.Connector */
{

	// Validate input parameters
	if ( ( theElementClient != null ) && ( theElementSupplier != null ) && ( theType != "" ) ) {

		// Cast the input values to objects so we get intellisense
		var curElementClient     as EA.Element;
		var curElementSupplier   as EA.Element;
		var curElementConnectors as EA.Collection;
		var curConnector         as EA.Connector;

		curElementClient         = theElementClient;
		curElementSupplier       = theElementSupplier;
		curConnector             = null;

		const theConnectorDirection = "Unspecified";

		// Check all Connectors in curElementClient whether the requested Connector already exists
		if ( skipDuplicate ) {
			curConnector         = CONGetElementConnectorByData( curElementClient, curElementClient.ElementID, curElementSupplier.ElementID, theType );
		}

		// If curConnector is not found, create a new Connector between curElementClient and curElementSupplier
		if ( curConnector == null )
		{

			curElementConnectors = curElementClient.Connectors;
			curConnector         = curElementConnectors.AddNew( curElementClient.Name, theType );

			// If curConnector is added, set the attributes
			if ( curConnector != null )
			{
				curConnector.Name       = "";     // Reset the dummy name as needed for AddNew
				curConnector.ClientID   = curElementClient.ElementID;
				curConnector.SupplierID = curElementSupplier.ElementID;
				curConnector.Direction  = theConnectorDirection;
				curConnector.Type       = theType;

				// Commit changes to the Repository
				curConnector.Update();
				curElementConnectors.Refresh();
				curElementClient.Update();
				curElementSupplier.Update();

			} else {
				LOGError("CONSetElementConnector could NOT create new Connector between Client(" + curElementClient.ElementID + ") and Supplier(" + curElementSupplier.ElementID + ") for Type " + theType + "!!!" );
				return null;
			}
		} else {
			LOGError("CONSetElementConnector skipped create duplicate Connector between Client(" + curElementClient.ElementID + ") and Supplier(" + curElementSupplier.ElementID + ") for Type " + theType + "!!!" );
			return null;
		}

		// Return the curConnector found or added
		return curConnector;

	}
}

/**
 * Deletes the specified Connector on the provided element.
 * NOTE: A Connector is found if the ClientID AND SupplierID and Type are identical.
 *
 * @param[in] theConnectorGUID (String) The GUID of the Connector to be deleted
 */
function CONDeleteConnectorByGUID( theConnectorGUID /* : String */ ) /* : void */
{

	// Cast theConnector to EA.Connector so we get intellisense
	var curConnector       as EA.Connector;
	var curElementClient   as EA.Element;
	var curElementSupplier as EA.Element;

	// Find the curConnectorGUID to identify the Connector
	let curConnectorGUID = theConnectorGUID;
	if ( curConnectorGUID == null ) {
		// Session.Output( "CONDeleteConnectorByGUID could NOT find curConnectorGUID so NOT deleted!!!");
		return "CONDeleteConnectorByGUID could NOT find curConnectorGUID so NOT deleted!!!";
	}

	// Find the curConnectorGUID
	curConnector = GetConnectorByGuid( curConnectorGUID );
	if ( curConnector == null ) {
		// Session.Output( "CONDeleteConnectorByGUID( " + curConnectorGUID + " ) could NOT find curConnector so NOT deleted!!!");
		return "CONDeleteConnectorByGUID( " + curConnectorGUID + " ) could NOT find curConnector so NOT deleted!!!";
	}


	// Find and process the source element defined by Start_Object_ID
	let curElementClientID = curConnector.ClientID;
	curElementClient       = GetElementByID( curElementClientID );
	if ( curElementClient == null ) {
		// Session.Output("CONDeleteConnectorByGUID( " + curConnectorGUID + " ) could NOT find curElementClient with curElementClientID = " + curElementClientID + " so NOT deleted!!!" );
		return "CONDeleteConnectorByGUID( " + curConnectorGUID + " ) could NOT find curElementClient with curElementClientID = " + curElementClientID + " so NOT deleted!!!";
	}

	// Find and process the target element defined by End_Object_ID
	let curElementSupplierID = curConnector.SupplierID;
	curElementSupplier       = GetElementByID( curElementSupplierID );
	if ( curElementSupplier == null ) {
		// Session.Output("CONDeleteConnectorByGUID( " + curConnectorGUID + " ) could NOT find curElementSupplier with curElementSupplierID = " + curElementSupplierID + " so NOT deleted!!!" );
		return "CONDeleteConnectorByGUID( " + curConnectorGUID + " ) could NOT find curElementSupplier with curElementSupplierID = " + curElementSupplierID + " so NOT deleted!!!";
	}

	// Process the Connector found by curConnectorGUID between curElementClient and curElementSupplier
	// Session.Output("CONDeleteConnectorByGUID( " + curConnectorGUID + " ) found curConnector.ConnectorID = " + curConnector.ConnectorID + ", ClientID = " + curConnector.ClientID + ", SupplierID = " + curConnector.SupplierID + "!!!" );
	// Delete the element as part of the curConnector.ClientID
	var curTempConnector   as EA.Connector;
	let curConnectorDeleted = false;
	// Find the index in the curElementClient.Connectors for the curConnector to delete
	for ( let i = 0 ; i < curElementClient.Connectors.Count ; i++ ) {
		curTempConnector = curElementClient.Connectors.GetAt( i );
		// Session.Output("CONDeleteConnectorByGUID TESTING curElementClient(" + i + ") where curTempConnector.ConnectorID = " + curTempConnector.ConnectorID + "!!!" );
		if ( curTempConnector.ConnectorID == curConnector.ConnectorID ) {
			curElementClient.Connectors.DeleteAt( i, false );
			// Session.Output("CONDeleteConnectorByGUID deleted curElementClient(" + i + ") where curConnector.ConnectorID = " + curConnector.ConnectorID + "!!!" );
			curConnectorDeleted = true;
			break; // Stop processing the rest of the Connectors in the for loop
		}
	}

	// Check curConnectorDeleted to commit updates to refresh changes
	if ( curConnectorDeleted ) {
		curElementClient.Connectors.Refresh();
		curElementSupplier.Connectors.Refresh();
	} else {
		// Session.Output("CONDeleteConnectorByGUID( " + curConnectorGUID + " ) could NOT find curElementSupplier with curElementSupplierID = " + curElementSupplierID + " so NOT deleted!!!" );
		return "CONDeleteConnectorByGUID( " + curConnectorGUID + " ) could NOT find curConnector.ConnectorID = " + curConnector.ConnectorID + " within " + curElementClient.Connectors.Count + " curElementClient.Connectors so NOT deleted!!!";
	}

	return "";
}
