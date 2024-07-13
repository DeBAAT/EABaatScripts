//[group=BaatPackageScripts]
!INC Local Scripts.EAConstants-JavaScript

/*
 * This code has been included from the default Project Browser template.
 * 
 * Script Name:	ModelCountDuplicateTaggedValues
 * Author:      J de Baat
 * Purpose:     Count all duplicate TaggedValues present in this model
 * 				NOTE: A TaggedValue is a duplicate if the Name AND Value are identical
 * Date:        19-08-2023
 * 
 */

/*
 * Get all duplicate values from the t_objectproperties table
 */
function GetDuplicateTaggedValues()
{

	// Get all the Object_IDs for which there are duplicates
	var strSQLQuery = "select op1.Object_ID, op1.Property from t_objectproperties op1"
                      + " where exists (select 1 from t_objectproperties op2 "
                      +						" where op1.Property   = op2.Property   "
                      +						"   and op1.Object_ID  = op2.Object_ID  "
                      +						"   and op1.Value      = op2.Value      "
                      +						"   and op1.PropertyID < op2.PropertyID "
                      +						" )"
                      + " order by op1.Object_ID; ";
	var sqlResponse = Repository.SQLQuery( strSQLQuery );
	Session.Output("strSQLQuery found sqlResponse= " + sqlResponse + "!!!" );

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
 * Project Browser Script main function
 */
function ModelCountDuplicateTaggedValues()
{
	// Get the type of element selected in the Project Browser
	var treeSelectedType = Repository.GetTreeSelectedItemType();
	var curElement as EA.Element;
	var strDuplicateTaggedValues = "";
	var intDuplicateTaggedValues = 8;

	Session.Output("+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++" );

	strDuplicateTaggedValues = GetDuplicateTaggedValues();

	// Process the strDuplicateTaggedValues found
	if ( strDuplicateTaggedValues.length > 0 ) {

		Session.Output("Found " + strDuplicateTaggedValues.length + " duplicates:" );

		if ( strDuplicateTaggedValues.length <= 10 ) {
			intDuplicateTaggedValues = strDuplicateTaggedValues.length;
		}

		// Show info found for max intDuplicateTaggedValues
		for ( var i = 0 ; i < intDuplicateTaggedValues ; i++ ) {
			curElement = Repository.GetElementByID( strDuplicateTaggedValues.item(i).text );
			Session.Output("==> duplicate curElement[" + strDuplicateTaggedValues.item(i).text + "].Name= " + curElement.Name + "!!!" );
		}

		Session.Prompt( "Found " + strDuplicateTaggedValues.length + " duplicates.", promptOK);
	} else {
		Session.Output("Found NO valid strDuplicateTaggedValues!!!" );
		Session.Prompt( "Found NO valid strDuplicateTaggedValues!!!", promptOK);
	}

	Session.Output("===========================================================================================" );

}

ModelCountDuplicateTaggedValues();
