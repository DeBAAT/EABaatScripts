//[group=BaatPackageScripts]
!INC Local Scripts.EAConstants-JavaScript

/*
 * This code has been included from the default Project Browser template.
 * 
 * Script Name:	ModelDeleteDuplicateTaggedValues
 * Author:      J de Baat
 * Purpose:     Delete all duplicate TaggedValues present in this model
 * 				NOTE: A TaggedValue is a duplicate if the Name, Value AND Object_ID are identical
 * Date:        20-08-2023
 * 
 */

/*
 * Get all duplicate values from the t_objectproperties table
 */
function GetDuplicateTaggedValues()
{

	// Get all the Object_IDs for which there are duplicates
	var strSQLQuery = "select op1.Object_ID, op1.Property from t_objectproperties op1"
                      + " where exists (select 1 from t_objectproperties op2"
                      +						" where op1.Property   = op2.Property"
                      +						"   and op1.Object_ID  = op2.Object_ID"
                      +						"   and op1.Value      = op2.Value"
                      +						"   and op1.PropertyID < op2.PropertyID"
                      +						" )"
                      + " group by op1.Object_ID"
                      + " order by op1.Object_ID"
                      + "; ";
	var sqlResponse = Repository.SQLQuery( strSQLQuery );
	// Session.Output("strSQLQuery found sqlResponse= " + sqlResponse + "!!!" );

	// Convert the sqlResponse from XML to an array of Object_IDs
	var arrResponse = convertXMLtoTagNameArray( sqlResponse, "Object_ID" );

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
 * Test whether the TaggedValue[ theTaggedValueIndex ] is a Duplicate in theElement
 *
 * Return:	false	theTaggedValueIndex too large OR theTaggedValueIndex NOT a duplicate
 * 			true	theTaggedValueIndex IS a duplicate
 */
function IsTaggedValueDuplicate( theElement, theTaggedValueIndex )
{

	// Cast the input values to objects so we get intellisense
	var curElement      as EA.Element;
	var curElementTags  as EA.Collection;
	var curTaggedValue  as EA.TaggedValue;
	var testTaggedValue as EA.TaggedValue;

	curElement     = theElement;
	curElementTags = curElement.TaggedValues;

	// Check theTaggedValueIndex
	if ( curElementTags.Count < theTaggedValueIndex )
	{
		// Clean up memory
		curElementTags = null;
		return false;
	}

	// Get the curTaggedValue to test against
	curTaggedValue = curElementTags.GetAt( theTaggedValueIndex );

	// Check all previous element tags for duplicate
	for ( var i = 0 ; i < theTaggedValueIndex ; i++ )
	{
		testTaggedValue = curElementTags.GetAt( i );
		if ( (curTaggedValue.Name == testTaggedValue.Name) && (curTaggedValue.Value == testTaggedValue.Value) ) 
		{
			// Clean up memory
			curElementTags = null;
			return true;
		}
		//Session.Output("COMPARED curElement(" + curElement.Name + ")(" + curElementTags.Count + ") found curTaggedValue.Name= " + curTaggedValue.Name + ", Value= " + curTaggedValue.Value + ", PropertyGUID= " + curTaggedValue.PropertyGUID + "!!!" );
	}

	// Clean up memory
	curElementTags = null;

	return false;

}

/*
 * Test whether theTaggedValue is a Duplicate and remove it if it is
 */
function TaggedValueRemoveDuplicate( theElement, theTaggedValueIndex, theNumDuplicates )
{

	// Cast the input values to objects so we get intellisense
	var curElement     as EA.Element;
	var curElementTags as EA.Collection;
	var curNumDuplicates;

	curElement       = theElement;
	curElementTags   = curElement.TaggedValues;
	curNumDuplicates = theNumDuplicates;

	// Check if the TaggedValue is a duplicate
	if ( IsTaggedValueDuplicate( curElement, theTaggedValueIndex ) )
	{
		// TaggedValue is Duplicate so remove from the Collection
		curElementTags.Delete( theTaggedValueIndex );
		curElementTags.Refresh();

		// Clean up memory
		curElementTags = null;

		return curNumDuplicates + 1;
	}

	// Clean up memory
	curElementTags = null;

	return curNumDuplicates;

}

/*
 * Check the TaggedValues of theElement provided as parameter
 */
function ProcessElement( theElement, theNumDuplicates )
{

	// Cast theElement to EA.Element so we get intellisense
	var curElement as EA.Element;
	var curElementTaggedValues as EA.Collection;
	var curNumDuplicates;
	var newNumDuplicates;

	curElement             = theElement;
	curElementTaggedValues = curElement.TaggedValues;
	curNumDuplicates       = theNumDuplicates;

	// Check for curElementTaggedValues
	if ( curElementTaggedValues.Count > 0 ) {

		// Process all curElementTaggedValues found
		for ( var i = curElementTaggedValues.Count - 1 ; i >= 0 ; i-- ) {
			// Test whether theTaggedValue is a Duplicate and remove it if it is
			newNumDuplicates = TaggedValueRemoveDuplicate( curElement, i, curNumDuplicates );
			curNumDuplicates = newNumDuplicates;
			// Session.Output("curElement(" + curElement.Name + ")(" + curElementTaggedValues.Count + ") found TaggedValues[" + i + "]= " + curElementTaggedValues.GetAt(i).Name + "!!!" );
		}

	} else {
		Session.Output("curElement(" + curElement.Name + ") found NO valid curElementTaggedValues!!!" );
	}

	// Clean up memory
	curElementTaggedValues = null;

	return curNumDuplicates;
}

/*
 * Project Browser Script main function
 */
function ModelDeleteDuplicateTaggedValues()
{
	// Get the type of element selected in the Project Browser
	var treeSelectedType = Repository.GetTreeSelectedItemType();
	var curElement as EA.Element;
	var strDuplicateTaggedValues = "";
	var intDuplicateTaggedValues = 10;
	var curNumDuplicates = 0;
	var newNumDuplicates = 0;

	Session.Output("+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++" );

	strDuplicateTaggedValues = GetDuplicateTaggedValues();

	// Process the strDuplicateTaggedValues found
	if ( strDuplicateTaggedValues.length > 0 ) {

		Session.Output("Found " + strDuplicateTaggedValues.length + " Elements with duplicates:" );

		// if ( strDuplicateTaggedValues.length <= 10 ) {
			intDuplicateTaggedValues = strDuplicateTaggedValues.length;
		// }

		// Show info found for max intDuplicateTaggedValues
		for ( var i = 0 ; i < intDuplicateTaggedValues ; i++ ) {

			// Process the curElement found
			curElement = Repository.GetElementByID( strDuplicateTaggedValues.item(i).text );
			newNumDuplicates = ProcessElement( curElement, curNumDuplicates );
			curNumDuplicates = newNumDuplicates;
			Session.Output("Processed duplicate curElement[" + i + "][" + strDuplicateTaggedValues.item(i).text + "].Name= ( " + curElement.Name + " ) with TaggedValues.Count=" + curElement.TaggedValues.Count + " and curNumDuplicates= " + curNumDuplicates + "!!!" );
		}

		Session.Prompt( "Deleted " + curNumDuplicates + " Duplicates from " + strDuplicateTaggedValues.length + " Elements.", promptOK);
	} else {
		Session.Output("Found NO valid strDuplicateTaggedValues!!!" );
		Session.Prompt( "Found NO Duplicate TaggedValues!!!", promptOK);
	}

	Session.Output("===========================================================================================" );

}

ModelDeleteDuplicateTaggedValues();
