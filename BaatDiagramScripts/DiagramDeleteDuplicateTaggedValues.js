//[group=BaatDiagramScripts]
!INC Local Scripts.EAConstants-JavaScript

/*
 * This code has been included from the default Diagram Script template.
 *
 * For all Elements on the selected diagram:
 *   For all TaggedValues for the curElement
 *      Mark a TaggedValue if it is a duplicate
 *   For all TaggedValues for the curElement
 *      Remove a TaggedValue if it is a duplicate
 *   Refresh Diagram
 *
 * NOTE: A TaggedValue is a duplicate if the Name AND Value are identical
 *
 * Script Name: DiagramDeleteDuplicateTaggedValues
 * Author:      J de Baat
 * Purpose:     Delete the duplicate TaggedValues of all elements occuring on the selected diagram
 * Date:        17-08-2023
 */


var DUPLICATE_PREFIX = "__DUPLICATE__";

/*
 * Test whether theTaggedValue is a Duplicate and remove it if it is
 */
function TaggedValueRemoveDuplicate( theElement, theTaggedValueIndex )
{

	// Cast the input values to objects so we get intellisense
	var curElement     as EA.Element;
	var curTaggedValue as EA.TaggedValue;
	var curElementTags as EA.Collection;
	curElement     = theElement;
	curElementTags = curElement.TaggedValues;

	// Check theTaggedValueIndex
	if ( curElementTags.Count < theTaggedValueIndex )
	{
		return -2;
	}

	// Get the curTaggedValue to test against
	curTaggedValue = curElementTags.GetAt( theTaggedValueIndex );

	// Check if the TaggedValue has a DUPLICATE_PREFIX
	if ( curTaggedValue.Name.startsWith( DUPLICATE_PREFIX ) )
	{
		Session.Output("curElement(" + curElement.Name + ") deleted DUPLICATE_PREFIX for curTaggedValue.item(" + theTaggedValueIndex + ").Name= " + curTaggedValue.Name + ", .Value= " + curTaggedValue.Value + "!!!" );
		curElementTags.Delete( theTaggedValueIndex );
		curElementTags.Refresh();
		return 1;
	} else {
		Session.Output("curElement(" + curElement.Name + ") found NO DUPLICATE_PREFIX for curTaggedValue.item(" + theTaggedValueIndex + ").Name= " + curTaggedValue.Name + ", .Value= " + curTaggedValue.Value + "!!!" );
		return 0;
	}

	return -1;
	//return ( (inputElement.ObjectType == 4) && (inputElement.Subtype == 76) );

}

/*
 * Test whether theTaggedValue is a Duplicate in theElement
 * Return results:	-2	theTaggedValueIndex too large
 * 					-1	theTaggedValueIndex NOT a duplicate
 * 					i	theTaggedValueIndex IS a duplicate
 */
function TaggedValueCheckDuplicate( theElement, theTaggedValueIndex )
{

	// Cast the input values to objects so we get intellisense
	var curElement     as EA.Element;
	var curTaggedValue as EA.TaggedValue;
	var curElementTags as EA.Collection;
	curElement     = theElement;
	curElementTags = curElement.TaggedValues;

	// Check theTaggedValueIndex
	if ( curElementTags.Count < theTaggedValueIndex )
	{
		return -2;
	}

	// Get the curTaggedValue to test against
	curTaggedValue = curElementTags.GetAt( theTaggedValueIndex );

	// Check all previous element tags for duplicate
	for ( var i = 0 ; i < theTaggedValueIndex ; i++ )
	{
		var testTaggedValue as EA.TaggedValue;
		testTaggedValue = curElementTags.GetAt( i );
		if ( (curTaggedValue.Name == testTaggedValue.Name) && (curTaggedValue.Value == testTaggedValue.Value) ) 
		{
			testTaggedValue.Name = DUPLICATE_PREFIX + testTaggedValue.Name;
			testTaggedValue.Update();
			return i;
		}
		//Session.Output("COMPARED curElement(" + curElement.Name + ")(" + curElementTags.Count + ") found curTaggedValue.Name= " + curTaggedValue.Name + ", Value= " + curTaggedValue.Value + ", PropertyGUID= " + curTaggedValue.PropertyGUID + "!!!" );
	}

	return -1;

}

/*
 * Get all TaggedValues for theTaggedValueObjectID from the t_objectproperties table
 */
function GetTaggedValueWithParameter( theTaggedValueParameter )
{

	// Get the TaggedValues registered for theTaggedValueParameter
	var strSQLQuery = "select * from t_objectproperties"
                      + " where t_objectproperties.PropertyID = '" + theTaggedValueParameter + "'"
                      + " order by t_objectproperties.Property";
	var sqlResponse = Repository.SQLQuery( strSQLQuery );

	// Convert the sqlResponse from XML to an array of TaggedValues
	var arrResponse = convertXMLtoTagNameArray( sqlResponse, "Property" );

	return arrResponse;

}

/*
 * Get all TaggedValues for theTaggedValueObjectID from the t_objectproperties table
 */
function GetTaggedValuesForElement( theTaggedValueObjectID )
{

	// Get all the TaggedValues registered for theTaggedValueObjectID
	var strSQLQuery = "select * from t_objectproperties"
                      + " where t_objectproperties.Object_ID = '" + theTaggedValueObjectID + "'"
                      + " order by t_objectproperties.Property";
	var sqlResponse = Repository.SQLQuery( strSQLQuery );

	// Convert the sqlResponse from XML to an array of TaggedValues
	var arrResponse = convertXMLtoTagNameArray( sqlResponse, "PropertyID" );

	return arrResponse;

}

/*
 * Extract an array from the XML resultset of an SQLQuery based on the xmlTagName
 */
function convertXMLtoTagNameArray( xmlString, xmlTagName )
{

	var xmlDOM = new COMObject( "MSXML2.DOMDocument" );
	xmlDOM.validateOnParse = false;
	xmlDOM.async = false;
	if ( xmlDOM.loadXML( xmlString ) ){
		var nodeList = xmlDOM.documentElement.selectNodes( '//' + xmlTagName );
		if ( nodeList.length > 0 ) {
			return nodeList;
		}
	}

	return false;

}

/*
 * Check the TaggedValues of theElement provided as parameter
 */
function CheckTaggedValuesElement( theElement )
{

	// Cast theElement to EA.Element so we get intellisense
	var curElement as EA.Element;
	var curElementTags as EA.Collection;
	curElement = theElement;
	curElementTags = curElement.TaggedValues;

	// List all element tags
	for ( var i = 0 ; i < curElementTags.Count ; i++ )
	{
		var curTaggedValue as EA.TaggedValue;
		curTaggedValue = curElementTags.GetAt( i );
		strResult = TaggedValueCheckDuplicate( curElement, i );
		Session.Output("curElement(" + curElement.Name + ")(" + curElementTags.Count + ") found strResult= " + strResult + " for curTaggedValue.Name= " + curTaggedValue.Name + ", Value= " + curTaggedValue.Value + ", PropertyGUID= " + curTaggedValue.PropertyGUID + "!!!" );
	}

	curElementTags.Refresh();
	arrElementTaggedValues = GetTaggedValuesForElement( curElement.ElementID );

	// Process the arrElementTaggedValues found
	if ( arrElementTaggedValues.length > 0 ) {

		// Add all getLegendPropString found for all strLegendTaggedValues
		for ( var i = arrElementTaggedValues.length - 1 ; i >= 0 ; i-- ) {
			// Test the TaggedValues found
			strResult = TaggedValueRemoveDuplicate( curElement, i );
			// Session.Output("curElement(" + curElement.Name + ")(" + arrElementTaggedValues.length + ") found TaggedValues.item(" + i + ")= " + arrElementTaggedValues.item(i).text + ", strResult= " + strResult + "!!!" );
		}

	} else {
		Session.Output("curElement(" + curElement.Name + ") found NO valid arrElementTaggedValues!!!" );
	}

}

/*
 * Diagram Script main function
 */
function DiagramDeleteDuplicateTaggedValues()
{
	// Get a reference to the current diagram
	var currentDiagram as EA.Diagram;
	currentDiagram = Repository.GetCurrentDiagram();

	Session.Output("+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++" );

	if ( currentDiagram != null )
	{
		Session.Output("Selected Diagram(DiagramID: " + currentDiagram.DiagramID + ") Name= " + currentDiagram.Name );

		// Get a reference to any selected connector/objects
		var diagramObjects as EA.Collection;
		var currentElement as EA.Element;
		diagramObjects = currentDiagram.DiagramObjects;

		// Check whether this diagram has any objects in it
		if ( diagramObjects.Count > 0 )
		{
			Session.Output("Selected diagramObjects.Count: " + diagramObjects.Count );
			// One or more diagram objects are selected
			for ( var i = 0 ; i < diagramObjects.Count ; i++ )
			{
				Session.Output("..........................................................................................." );

				// Process the currentDiagramElement
				var currentDiagramElement as EA.Element;
				var currentElement as EA.Element;
				currentDiagramElement = diagramObjects.GetAt( i );
				currentElement = Repository.GetElementByID( currentDiagramElement.ElementID );

				// Check the TaggedValues of the currentElement
				CheckTaggedValuesElement( currentElement );
			}

			// Reload diagram when all processing is done
			Repository.ReloadDiagram( currentDiagram.DiagramID );
		}
		else
		{
			// No objects on this diagram
			Session.Output("No objects on this diagram" );
		}
	}
	else
	{
		Session.Prompt( "This script requires a diagram to be visible.", promptOK)
	}

	Session.Output("===========================================================================================" );

}

DiagramDeleteDuplicateTaggedValues();