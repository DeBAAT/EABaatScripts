//[group=BaatDiagramScripts]
!INC Local Scripts.EAConstants-JavaScript

/*
 * Script Name: RefreshDynamicLegend
 * Author:      J de Baat
 * Purpose:     Refresh the set of color definitions of all legend elements supporting all values occuring on the selected diagram
 * Date:        04-08-2024
 *
 * For all LegendElements on the selected diagram:
 *   If TaggedValue filter
 *      Find all TaggedValues for this Value
 *      Update t_xref list of properties for this legend
 *   Refresh Diagram
*/

/*
 * A list of colors to use for generating the list of legend values
 * Hexadecimal to Decimal converter via https://www.rapidtables.com/convert/number/hex-to-decimal.html
 */
var legendColors =[ "1029550",		// #0FB5AE
            "4212426",		// #4046CA
            "16155921",		// #F68511
            "14564738",		// #DE3D82
            "8291578",		// #7E84FA
            "7528554",		// #72E06A
            "1342195",		// #147AF3
            "7546579",		// #7326D3
            "15255040",		// #E8C600
            "13327616",		// #CB5D00
            "36701",		// #008F5D
            "12380465",		// #BCE931
            "0"				// #black
];

/*
 * Test whether an element is a Legend
 */
function ElementIsLegend( theElement )
{
	// Cast theElement to EA.Element so we get intellisense
	var inputElement as EA.Element;
	inputElement      = theElement;

	return ( (inputElement.ObjectType == 4) && (inputElement.Subtype == 76) );

}

/*
 * Get the filter value for the LegendElement
 */
function GetLegendFilter( theElement )
{
	// Cast theElement to EA.Element so we get intellisense
	var inputElement as EA.Element;
	inputElement      = theElement;

	// Find the TaggedValue in the string of format "*LegendTypeObj=Filter=" + "TaggedValue." + strLegendTaggedValue + ":;*"
	var strLegendTaggedValue01 = inputElement.StyleEx.split("LegendTypeObj=Filter=");
	var strLegendTaggedValue02 = strLegendTaggedValue01[1].split(":;");
	var strLegendTaggedValue03 = strLegendTaggedValue02[0].split(".");
	if ( strLegendTaggedValue03[0] == "TaggedValue" )
	{
		// Check for strLegendTaggedValue to be of format strLegendValue + ":AppliesTo=*"
		var strLegendTaggedValue04 = strLegendTaggedValue03[1].split(":AppliesTo=");
		return strLegendTaggedValue04[0];
	}
	return false;

}

/*
 * Get all values for theTaggedValueProperty from the t_objectproperties table
 */
function GetLegendTaggedValuesForProperty( theTaggedValueProperty )
{

	// Get all the TaggedValues registered for theTaggedValueProperty
	var strSQLQuery = "select distinct t_objectproperties.Value as TaggedValues from t_objectproperties"
                      + " where t_objectproperties.Property = '" + theTaggedValueProperty + "'"
                      + " order by t_objectproperties.Value";
	var sqlResponse = Repository.SQLQuery( strSQLQuery );

	// Convert the sqlResponse from XML to an array of TaggedValues
	var arrResponse = convertXMLtoTagNameArray( sqlResponse, "TaggedValues" );

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
 * Update the Description value for the LegendElement in the t_xref table
 */
function updateLegendElementXRef( theElement, theLegendDescription )
{

	// Cast theElement to EA.Element so we get intellisense
	var inputElement as EA.Element;
	inputElement = theElement;

	// Get all the TaggedValues registered for theTaggedValueProperty
	var strSQLQuery = "update t_xref set Description = '" + theLegendDescription + "'"
                      + " where Name = 'CustomProperties'"
                      + " and Type = 'element property'"
                      + " and Client = '" + inputElement.ElementGUID + "'";
	var sqlResponse = Repository.Execute( strSQLQuery );

	return true;

	// TODO: Check sqlResponse as result of Repository.Execute
	//Session.Output("strSQLQuery ==> Resulted in: " + sqlResponse );
	//return sqlResponse;

}

/*
 * Make a new property string for the theValueName and theValueIndex
 */
function getLegendPropString( theValueName, theValueIndex )
{

	var strPropOutput = "";

	strPropOutput += "@PROP=";
	strPropOutput += "@NAME=";
	strPropOutput += theValueName;
	strPropOutput += "@ENDNAME;";
	strPropOutput += "@TYPE=LEGEND_OBJECTSTYLE@ENDTYPE;";
	strPropOutput += "@VALU=#Back_Ground_Color#=";
	strPropOutput += legendColors[theValueIndex % legendColors.length ];
	strPropOutput += ";";
	strPropOutput += "#Pen_Color#=0;";
	strPropOutput += "#Pen_Size#=1;";
	strPropOutput += "#Legend_Type#=LEGEND_OBJECTSTYLE;";
	strPropOutput += "@ENDVALU;";
	strPropOutput += "@PRMT=";
	strPropOutput += theValueIndex;
	strPropOutput += "@ENDPRMT;";
	strPropOutput += "@ENDPROP;";

	return strPropOutput;

}

/*
 * Make a new property string for theLegendName and theLegendTitle
 */
function getLegendStyleSetting( theLegendName, theLegendTitle )
{

	var strPropOutput = "";

	strPropOutput += "@PROP=";
	strPropOutput += "@NAME=";
	strPropOutput += theLegendName;
	strPropOutput += "@ENDNAME;";
	strPropOutput += "@TYPE=LEGEND_STYLE_SETTINGS@ENDTYPE;";
	strPropOutput += "@VALU=#TITLE#=";
	strPropOutput += theLegendTitle;
	strPropOutput += ";";
	strPropOutput += "@ENDVALU;";
	strPropOutput += "@PRMT=";
	strPropOutput += "@ENDPRMT;";
	strPropOutput += "@ENDPROP;";

	return strPropOutput;

}

/*
 * Process a LegendElement provided as parameter
 */
function ProcessLegendElement( theElement )
{

	// Cast theElement to EA.Element so we get intellisense
	var inputElement as EA.Element;
	currentElement    = theElement;

	// Check whether the currentElement is a Legend
	if ( ElementIsLegend( currentElement ) )
	{
		var strLegendFilterProperty = "";
		var strLegendTaggedValues   = "";
		var strLegendProperty       = "";

		strLegendFilterProperty = GetLegendFilter( currentElement );
		strLegendTaggedValues   = GetLegendTaggedValuesForProperty( strLegendFilterProperty );

		// Process the strLegendTaggedValues found
		if ( strLegendTaggedValues.length > 0 ) {

			// Add all getLegendPropString found for all strLegendTaggedValues
			for ( var i = 0 ; i < strLegendTaggedValues.length ; i++ ) {
				strLegendProperty += getLegendPropString( strLegendTaggedValues.item(i).text, i );
			}

			// Add the getLegendStyleSetting found
			strLegendProperty += getLegendStyleSetting( currentElement.Name, currentElement.Name );

			// Refresh the currentElement and Update the currentDiagram to reflect the changes
			var boolResult = updateLegendElementXRef( currentElement, strLegendProperty );
			if ( boolResult ) {
				Session.Output("Legend(" + currentElement.Name + ") UPDATED successfully!!!" );
			} else {
				Session.Output("Legend(" + currentElement.Name + ") update NOT successful!!!" );
			}
		} else {
			Session.Output("Legend(" + currentElement.Name + ") found NO valid strLegendTaggedValues!!!" );
		}
	}

}

/*
 * Diagram Script main function
 */
function RefreshDynamicLegend()
{

	// Show the script output window
	Repository.EnsureOutputVisible( "Script" );

	Session.Output( "======================================= Started RefreshDynamicLegend " );

	// Get a reference to the current diagram
	var currentDiagram as EA.Diagram;
	currentDiagram = Repository.GetCurrentDiagram();

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
				// Process the currentDiagramElement
				var currentDiagramElement as EA.Element;
				var currentElement        as EA.Element;
				currentDiagramElement      = diagramObjects.GetAt( i );
				currentElement             = Repository.GetElementByID( currentDiagramElement.ElementID );

				// Process the currentElement when it is a Legend
				ProcessLegendElement( currentElement );
			}

			// Reload diagram when all processing is done
			Repository.ReloadDiagram( currentDiagram.DiagramID );
		}
		else
		{
			// No objects on this diagram
			Session.Output("No objects on this diagram to be processed." );
		}
	}
	else
	{
		Session.Prompt( "This script requires a diagram to be visible.", promptOK)
	}

	Session.Output( "======================================= Finished RefreshDynamicLegend " );

}

RefreshDynamicLegend();