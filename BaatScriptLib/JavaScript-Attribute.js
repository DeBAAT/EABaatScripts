//[group=BaatScriptLib]
!INC BaatScriptLib.BaatScript-Logging

/**
 * @file JavaScript-Attribute 
 * This script library contains helper functions for working with Attributes. Functions 
 * provided by this module are identified by the prefix CON.
 *
 * @author	J. de Baat, based on JavaScript-TaggedValue by Sparx Systems
 * @date	08-10-2025
 */

/**
 * Retrieves the Attribute object from the provided element whose data matches the specified parameters.
 * If the element does not exist, or does not contain a Attribute with the specified data, null
 * is returned.
 * NOTE: A Attribute is found if the ClientID AND SupplierID and Type are identical.
 *
 * @param[in] theElement (EA.Element) The element to retrieve the Attribute value from
 * @param[in] theClientID (String) The ID of the Element registered as Source of the Attribute
 * @param[in] theSupplierID (String) The ID of the Element registered as Target of the Attribute
 * @param[in] theType (String) The type of the Attribute to be found, if empty string, any type will match
 *
 * @return The object of the requested Attribute found, null when not found
 */
function CONGetElementAttributeByData( theElement /* : EA.Element */, theClientID /* : ID */, theSupplierID /* : ID */, theType /* : String */ ) /* : EA.Attribute */
{

	// Validate input parameters
	if ( ( theElement != null ) && ( theClientID > 0 ) ) {

		// Cast the input values to objects so we get intellisense
		var curElement           as EA.Element;
		var curElementAttributes as EA.Collection;
		var curAttribute         as EA.Attribute;

		try
		{

			curElement           = theElement;
			curElementAttributes = curElement.Attributes;

			// Check all element Attributes against data to find
			let curElementAttributesCount = curElementAttributes.Count;
			for ( var i = 0 ; i < curElementAttributesCount ; i++ )
			{
				curAttribute = curElementAttributes.GetAt( i );
				if ( ( curAttribute.ClientID   == theClientID ) &&
					 ( curAttribute.SupplierID == theSupplierID ) &&
					 ( ( "" === theType ) || ( curAttribute.Type == theType ) ) ) 
				{
					// Attribute found so clean up memory and return object
					curElementAttributes = null;
					return curAttribute;
				}
			}
		}
		catch(catch_err)
		{
			BLOGError("CONGetElementAttributeByData found Error: " + catch_err + "!!!" );
		}

		// Clean up memory
		curElementAttributes = null;
	}

	return null;

}

/**
 * Sets the specified Attribute on the provided element. If the provided element does not already
 * contain a Attribute with the specified data, a new Attribute is created.
 * If a Attribute already exists with the specified data then the action is ignored.
 * NOTE: A Attribute is found if the ClientID AND SupplierID and Type are identical.
 *
 * @param[in] theElementClient (EA.Element) The element to add the Attribute to and registered as Source of the Attribute
 * @param[in] theElementSupplier (EA.Element) The element to be registered as Target of the Attribute
 * @param[in] theType (String) The type of the Attribute to be found, if empty string, any type will match
 * @param[in] skipDuplicate (boolean) If set to true, check for existing connector to prevent duplicate
 *
 * @return The object of the Attribute added or found, null in case of error
 */
function CONSetElementAttribute( theElementClient /* : EA.Element */, theElementSupplier /* : EA.Element */, theType /* : String */, skipDuplicate /* : Boolean */ ) /* : EA.Attribute */
{

	// Validate input parameters
	if ( ( theElementClient != null ) && ( theElementSupplier != null ) && ( theType != "" ) ) {

		// Cast the input values to objects so we get intellisense
		var curElementClient     as EA.Element;
		var curElementSupplier   as EA.Element;
		var curElementAttributes as EA.Collection;
		var curAttribute         as EA.Attribute;

		const theAttributeDirection = "Unspecified";

		try
		{

			curElementClient   = theElementClient;
			curElementSupplier = theElementSupplier;
			curAttribute       = null;

			// Check all Attributes in curElementClient whether the requested Attribute already exists
			if ( skipDuplicate ) {
				curAttribute   = CONGetElementAttributeByData( curElementClient, curElementClient.ElementID, curElementSupplier.ElementID, theType );
			}

			BLOGTrace("CONSetElementAttribute testing Attribute between Client(" + curElementClient.ElementID + ") and Supplier(" + curElementSupplier.ElementID + ") for Type " + theType + "!!!" );

			// If curAttribute is not found, create a new Attribute between curElementClient and curElementSupplier
			if ( curAttribute == null )
			{

				curElementAttributes = curElementClient.Attributes;
				curAttribute         = curElementAttributes.AddNew( curElementClient.Name, theType );
				curAttribute.Update();
				curElementAttributes.Refresh();

				// If curAttribute is added, set the attributes
				if ( curAttribute != null )
				{
					curAttribute.Name       = "";     // Reset the dummy name as needed for AddNew
					curAttribute.ClientID   = curElementClient.ElementID;
					curAttribute.SupplierID = curElementSupplier.ElementID;
					curAttribute.Direction  = theAttributeDirection;
					curAttribute.Type       = theType;

					// Commit changes to the Repository
					curAttribute.Update();
					curElementAttributes.Refresh();
					curElementClient.Update();
					curElementSupplier.Update();
					BLOGDebug("CONSetElementAttribute created new Attribute(" + curAttribute.AttributeID + ") between Client(" + curElementClient.ElementID + ") and Supplier(" + curElementSupplier.ElementID + ") for Type " + theType + "!!!" );
					return curAttribute;

				} else {
					BLOGError("CONSetElementAttribute could NOT create new Attribute between Client(" + curElementClient.ElementID + ") and Supplier(" + curElementSupplier.ElementID + ") for Type " + theType + "!!!" );
					return null;
				}
			} else {
				BLOGDebug("CONSetElementAttribute skipped create duplicate Attribute between Client(" + curElementClient.ElementID + ") and Supplier(" + curElementSupplier.ElementID + ") for Type " + theType + "!!!" );
				return curAttribute;
			}

			// Return the curAttribute found or added
			BLOGDebug("CONSetElementAttribute returns found Attribute(" + curAttribute.AttributeID + ") between Client(" + curElementClient.ElementID + ") and Supplier(" + curElementSupplier.ElementID + ") for Type " + theType + "!!!" );
			return curAttribute;
		}
		catch(catch_err)
		{
			BLOGError("CONSetElementAttribute found Error: " + catch_err + "!!!" );
			return null;
		}

	}

	return null;

}

/**
 * Deletes the specified Attribute on the provided element.
 * NOTE: A Attribute is found if the ClientID AND SupplierID and Type are identical.
 *
 * @param[in] theAttributeGUID (String) The GUID of the Attribute to be deleted
 */
function CONDeleteAttributeByGUID( theAttributeGUID /* : String */ ) /* : void */
{

	// Cast theAttribute to EA.Attribute so we get intellisense
	var curAttribute       as EA.Attribute;
	var curElementClient   as EA.Element;
	var curElementSupplier as EA.Element;

	try
	{

		// Find the curAttributeGUID to identify the Attribute
		let curAttributeGUID = theAttributeGUID;
		if ( curAttributeGUID == null ) {
			// BLOGTrace( "CONDeleteAttributeByGUID could NOT find curAttributeGUID so NOT deleted!!!");
			return "CONDeleteAttributeByGUID could NOT find curAttributeGUID so NOT deleted!!!";
		}

		// Find the curAttributeGUID
		curAttribute = GetAttributeByGuid( curAttributeGUID );
		if ( curAttribute == null ) {
			// BLOGTrace( "CONDeleteAttributeByGUID( " + curAttributeGUID + " ) could NOT find curAttribute so NOT deleted!!!");
			return "CONDeleteAttributeByGUID( " + curAttributeGUID + " ) could NOT find curAttribute so NOT deleted!!!";
		}


		// Find and process the source element defined by Start_Object_ID
		let curElementClientID = curAttribute.ClientID;
		curElementClient       = GetElementByID( curElementClientID );
		if ( curElementClient == null ) {
			// BLOGTrace("CONDeleteAttributeByGUID( " + curAttributeGUID + " ) could NOT find curElementClient with curElementClientID = " + curElementClientID + " so NOT deleted!!!" );
			return "CONDeleteAttributeByGUID( " + curAttributeGUID + " ) could NOT find curElementClient with curElementClientID = " + curElementClientID + " so NOT deleted!!!";
		}

		// Find and process the target element defined by End_Object_ID
		let curElementSupplierID = curAttribute.SupplierID;
		curElementSupplier       = GetElementByID( curElementSupplierID );
		if ( curElementSupplier == null ) {
			// BLOGTrace("CONDeleteAttributeByGUID( " + curAttributeGUID + " ) could NOT find curElementSupplier with curElementSupplierID = " + curElementSupplierID + " so NOT deleted!!!" );
			return "CONDeleteAttributeByGUID( " + curAttributeGUID + " ) could NOT find curElementSupplier with curElementSupplierID = " + curElementSupplierID + " so NOT deleted!!!";
		}

		// Process the Attribute found by curAttributeGUID between curElementClient and curElementSupplier
		// BLOGTrace("CONDeleteAttributeByGUID( " + curAttributeGUID + " ) found curAttribute.AttributeID = " + curAttribute.AttributeID + ", ClientID = " + curAttribute.ClientID + ", SupplierID = " + curAttribute.SupplierID + "!!!" );
		// Delete the element as part of the curAttribute.ClientID
		var curTempAttribute   as EA.Attribute;
		let curAttributeDeleted = false;
		// Find the index in the curElementClient.Attributes for the curAttribute to delete
		let curElementClientAttributesCount = curElementClient.Attributes.Count;
		for ( let i = 0 ; i < curElementClientAttributesCount ; i++ ) {
			curTempAttribute = curElementClient.Attributes.GetAt( i );
			// BLOGTrace("CONDeleteAttributeByGUID TESTING curElementClient(" + i + ") where curTempAttribute.AttributeID = " + curTempAttribute.AttributeID + "!!!" );
			if ( curTempAttribute.AttributeID == curAttribute.AttributeID ) {
				curElementClient.Attributes.DeleteAt( i, false );
				// BLOGTrace("CONDeleteAttributeByGUID deleted curElementClient(" + i + ") where curAttribute.AttributeID = " + curAttribute.AttributeID + "!!!" );
				curAttributeDeleted = true;
				break; // Stop processing the rest of the Attributes in the for loop
			}
		}

		// Check curAttributeDeleted to commit updates to refresh changes
		if ( curAttributeDeleted ) {
			curElementClient.Attributes.Refresh();
			curElementSupplier.Attributes.Refresh();
		} else {
			// BLOGTrace("CONDeleteAttributeByGUID( " + curAttributeGUID + " ) could NOT find curElementSupplier with curElementSupplierID = " + curElementSupplierID + " so NOT deleted!!!" );
			return "CONDeleteAttributeByGUID( " + curAttributeGUID + " ) could NOT find curAttribute.AttributeID = " + curAttribute.AttributeID + " within " + curElementClient.Attributes.Count + " curElementClient.Attributes so NOT deleted!!!";
		}
	}
	catch(catch_err)
	{
		BLOGError("CONDeleteAttributeByGUID found Error: " + catch_err + "!!!" );
	}

	return "";
}
