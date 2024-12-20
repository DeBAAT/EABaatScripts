//[group=BaatScriptLib]
!INC Local Scripts.EAConstants-JavaScript
!INC EAScriptLib.JavaScript-Dialog
!INC BaatScriptLib.BaatScript-Logging

/*
 * Script Name:	DiagramAnalysisLayoutCommon
 * Author:		J de Baat
 * Purpose:		Dynamically draw elements on the diagram as indicated by the TaggedValues defined in the selected element
 * Date:		20-12-2024
 *
 * This script takes a number of selected elements on a diagram to start the analysis.
 * For each selected element on a diagram
 *   If the Keywords property of the element contains the DALKeywordsTag
 *     Collect all referenced elements to a set for the first level
 *     Process the current level for all elements in the set:
 *       For each referenced element in the set
 *         If element not yet drawn on the diagram
 *           Draw the element on the indicated location of the diagram
 *           Collect all referenced elements to a set for the next level
 *       Recursively process the set of elements found
 *
 */

// The level to log at
var   BLOGLEVEL = BLOGLEVEL_WARNING;		// Choose from: BLOGLEVEL_ERROR, BLOGLEVEL_WARNING, BLOGLEVEL_INFO, BLOGLEVEL_DEBUG, BLOGLEVEL_TRACE

/*
 * mapDiagramLayoutValues is a map of [Key,Value] to use when generating the analysis diagram
 * Each mapDiagramLayoutValue has a default defined which can be replaced with a value defined as TaggedValue for the selected element
 */
var mapDiagramLayoutValues;

/*
 * mapDALStageLevels is a map of [Key,Value] which contains the number of tables for that particular stage level
 */
var mapDALStageLevels = null;

var objGlobalDALDiagram as EA.Diagram;

const DALDirection     = "DALDirection";			// Whether to analyse in Horizontal (default) or Vertical Direction
const DALNumLevels     = "DALNumLevels";			// Number of levels to analyse
const DALFilter        = "DALFilter";				// Connector TaggedValue to filter what to show
const DALSpaceHor      = "DALSpaceHor";				// Horizontal spacing between elements
const DALSpaceVer      = "DALSpaceVer";				// Vertical spacing between elements
const DALStartHor      = "DALStartHor";				// Horizontal start for elements
const DALStartVer      = "DALStartVer";				// Vertical start for elements
const DALElementWidth  = "DALElementWidth";			// Width of an element to draw
const DALElementHeight = "DALElementHeight";		// Height of an element to draw

const DALTagPrefix     = "DAL";						// String prefix to indicate that this TaggedValue can be used for DiagramAnalysisLayout
const DALFilterNone    = "DALFilterNone";			// Default value to indicate not to use any DALFilter
const DALKeywordsTag   = "diagramanalysislayout";	// String Tag to indicate that this element can be used for DiagramAnalysisLayout
const DALHorizontal    = "Horizontal";				// Analyse in Horizontal (default) Direction
const DALVertical      = "Vertical";				// Analyse in Vertical Direction

const DALTaggedValueStage        = "Stage";			// Definition of the TV Stage property
const DALTaggedValueStageDefault = "DefaultStage";	// Definition of the TV Stage property default value if not defined

/*
 * Create the set of defaults for mapDiagramLayoutValues
 */
function DALCreateDiagramLayoutValues()
{

	try
	{

		// Clean up memory and create a new Map
		mapDiagramLayoutValues = null;
		mapDiagramLayoutValues = new Map();

		mapDiagramLayoutValues.set( DALDirection,     DALHorizontal );
		mapDiagramLayoutValues.set( DALFilter,        DALFilterNone );
		mapDiagramLayoutValues.set( DALNumLevels,     4             );
		mapDiagramLayoutValues.set( DALSpaceHor,      20            );
		mapDiagramLayoutValues.set( DALSpaceVer,      20            );
		mapDiagramLayoutValues.set( DALStartHor,      50            );
		mapDiagramLayoutValues.set( DALStartVer,      100           );
		mapDiagramLayoutValues.set( DALElementWidth,  150           );
		mapDiagramLayoutValues.set( DALElementHeight, 60            );
	}
	catch (catch_err)
	{
		BLOGError("DALCreateDiagramLayoutValues found Error: " + catch_err.message + "!!!" );
	}

}

/*
 * Get the set of values for mapDiagramLayoutValues from the TaggedValues of theElement
 */
function DALGetValuesFromElement( theElement )
{

	// Cast the input values to objects so we get intellisense
	var curElement      as EA.Element;
	var curElementTags  as EA.Collection;
	var curTaggedValue  as EA.TaggedValue;

	try
	{

		curElement     = theElement;
		curElementTags = curElement.TaggedValues;

		// Check all element tags for mapDiagramLayoutValues (starting with "DAL")
		let curElementTagsCount = curElementTags.Count;
		for ( var i = 0 ; i < curElementTagsCount ; i++ )
		{
			curTaggedValue = curElementTags.GetAt( i );
			if ( (curTaggedValue.Name.startsWith( DALTagPrefix )) ) 
			{
				// Get the mapDiagramLayoutValues from the curTaggedValue
				mapDiagramLayoutValues.set( curTaggedValue.Name,   curTaggedValue.Value );
				BLOGTrace("Found curTaggedValue to use for mapDiagramLayoutValues[ " + curTaggedValue.Name + " ] = " + curTaggedValue.Value + " !!!" );
			}
		}
	}
	catch (catch_err)
	{
		BLOGError("DALGetValuesFromElement found Error: " + catch_err.message + "!!!" );
	}

	// Clean up memory
	curElementTags = null;

	return false;

}

/*
 * Get the location of an element based on the parameters
 */
function DALGetElementLocation( numX, numY )
{

	try
	{

		// Calculate the Offset values
		let intOffsetX = numX * ( Number( mapDiagramLayoutValues.get( DALSpaceHor ) ) + Number( mapDiagramLayoutValues.get( DALElementWidth  ) ) );
		let intOffsetY = numY * ( Number( mapDiagramLayoutValues.get( DALSpaceVer ) ) + Number( mapDiagramLayoutValues.get( DALElementHeight ) ) );

		// Create new string from information calculated
		const locL = Number( intOffsetX ) + Number( mapDiagramLayoutValues.get( DALStartHor      ) );
		const locR = Number( locL       ) + Number( mapDiagramLayoutValues.get( DALElementWidth  ) );
		const locT = Number( intOffsetY ) + Number( mapDiagramLayoutValues.get( DALStartVer      ) );
		const locB = Number( locT       ) + Number( mapDiagramLayoutValues.get( DALElementHeight ) );
		const strElementLocation = "l=" + locL + ";r=" + locR + ";t=" + locT + ";b=" + locB + ";";

		return strElementLocation;
	}
	catch (catch_err)
	{
		BLOGError("DALGetElementLocation found Error: " + catch_err.message + "!!!" );
	    return "";
	}

}

/*
 * An element is for DiagramAnalysisLayout when the KeywordsTag contains the string defined by DALKeywordsTag
 */
function DALIsDiagramAnalysisLayoutElement( theElement )
{
	// Cast theElement to EA.Element so we get intellisense
	var curElement as EA.Element;
	curElement      = theElement;

	try
	{
		var strKeywordsTag = curElement.Tag.toLowerCase();
		var idxKeywordsTag = strKeywordsTag.indexOf( DALKeywordsTag );
		return ( idxKeywordsTag >= 0 );
	}
	catch (catch_err)
	{
		BLOGError("DALIsDiagramAnalysisLayoutElement found Error: " + catch_err.message + "!!!" );
	    return false;
	}

    return false;

}

/*
 * Test whether an element is a Legend
 */
function DALElementIsLegend( theElement )
{
	// Cast theElement to EA.Element so we get intellisense
	var inputElement as EA.Element;
	inputElement      = theElement;

	return ( (inputElement.ObjectType == 4) && (inputElement.Subtype == 76) );

}

/*
 * Show the element for DiagramAnalysisLayout when theConnector DALFilter TaggedValue contains the string defined by strDALFilterValue
 */
function DALShowConnector( theConnector )
{
	// Cast theConnector to EA.Connector so we get intellisense
	var curConnector     as EA.Connector;
	var curConnectorTag  as EA.TaggedValue;
	var curConnectorTags as EA.Collection;
	curConnector          = theConnector;

	// Only check when DALFilter defined
	if ( mapDiagramLayoutValues.get( DALFilter ) === DALFilterNone ) {
		return true;
	}

	try
	{

		let strDALFilterValue = mapDiagramLayoutValues.get( DALFilter );
		strDALFilterValue = strDALFilterValue.toLowerCase();

		// Process all curConnectorTags in curConnector
		curConnectorTags = curConnector.TaggedValues;
		let curConnectorTagsCount = curConnectorTags.Count;
		for ( let i = 0 ; i < curConnectorTagsCount ; i++ )
		{
			curConnectorTag  = curConnectorTags.GetAt( i );

			// Test whether value of DALFilter contains the element needed
			if ( (curConnectorTag.Name.startsWith( DALFilter )) ) 
			{
				var curConnectorTagFilter = curConnectorTag.Value.toLowerCase();
				var idxDALFilterValue = curConnectorTagFilter.indexOf( strDALFilterValue );
				if ( idxDALFilterValue >= 0 ) {
					BLOGTrace("DALShowConnector Processed curConnectorTags(" + i + ") of [" + curConnectorTagsCount + "] and curConnectorTag.Name = " + curConnectorTag.Name + "!!!" );
					return true;
				}
			}
		}
		BLOGTrace("DALShowConnector Processed " + curConnectorTagsCount + " curConnectorTags for curConnector.ConnectorID = " + curConnector.ConnectorID + "!!!" );
	}
	catch (catch_err)
	{
		BLOGError("DALShowConnector found Error: " + catch_err.message + "!!!" );
	    return false;
	}

	BLOGDebug("DALShowConnector Processed " + curConnectorTags.Count + " curConnectorTags for curConnector.ConnectorID = " + curConnector.ConnectorID + "!!!" );
    return false;

}

/*
 * Get the requested Element from theCollectionObjects using theElementID as parameter
 */
function DALGetCollectionObjectByID( theCollectionObjects, theElementID )
{
	// Cast theElement to EA.Element so we get intellisense
	var curCollectionObjects as EA.Collection;
	var curCollectionObject  as EA.Element;

	// Check all Elements in theCollectionObjects whether the requested curElement is already defined
	curCollectionObjects = theCollectionObjects;

	try
	{

		// Loop over curCollectionObjects to find theElementID
		let curCollectionObjectsCount = curCollectionObjects.Count;
		for ( var i = 0 ; i < curCollectionObjectsCount ; i++ )
		{
			curCollectionObject = curCollectionObjects.GetAt( i );
			if ( curCollectionObject.ElementID === theElementID ) {
				BLOGTrace( "DALGetCollectionObjectByID found CollectionObject ( " + curCollectionObject.ElementID + " ) as part of " + curCollectionObjectsCount + " CollectionObjects!!!" );
				return curCollectionObject;
			}
			BLOGTrace( "DALGetCollectionObjectByID TESTED CollectionObject ( " + curCollectionObject.ElementID + " ) against theElementID " + theElementID + " !!!" );
		}
	}
	catch (catch_err)
	{
		BLOGError("DALGetCollectionObjectByID found Error: " + catch_err.message + "!!!" );
	    return null;
	}

	// theElementName not found as part of curCollectionObjects
	BLOGDebug( "DALGetCollectionObjectByID DID NOT FIND theElementID ( " + theElementID + " ) as part of CollectionObjects!!!" );
	return null;

}

/*
 * Get the index of theKey contained in theMap
 */
function DALGetMapIndex( theMap, theKey )
{

	try
	{
		let index = 0;
		for (const mapKey of theMap.keys()) {
			// Return the index if the key matches
			if ( mapKey == theKey ) {
				return index;
			}
			index++;
		}
		// Return null if theKey is not found
		return 0;
	}
	catch (catch_err)
	{
		BLOGError("DALGetMapIndex found Error: " + catch_err.message + "!!!" );
	    return 0;
	}

}

/*
 * Get the requested Element from theCollectionObjects using theElementID as parameter
 */
function DALGetElementNewName( theElement )
{
	// Cast theElement to EA.Element so we get intellisense
	var curElement  as EA.Element;
	curElement       = theElement;

	try
	{
		var curElementStageValue  = DALTaggedValueStageDefault;
		var curElementStageObject = curElement.TaggedValues.GetByName( DALTaggedValueStage );
		if ( curElementStageObject != undefined ) {
			curElementStageValue = curElementStageObject.Value;
		}
		BLOGDebug( "DALGetElementNewName found curElementStageValue( " + curElementStageValue + " ) for curElement( " + curElement.Name + " )!!!" );

		// Get the current value for mapDALStageLevels if it exists
		var curElementStageNumbers = 0;
		if ( mapDALStageLevels.has( curElementStageValue ) ) {
			curElementStageNumbers = Number( mapDALStageLevels.get( curElementStageValue ) ) + 1;
		}
		mapDALStageLevels.set( curElementStageValue, curElementStageNumbers );
		var curElementStageIndex = DALGetMapIndex( mapDALStageLevels, curElementStageValue );
		BLOGDebug( "DALGetElementNewName found curElementStageNum[ " + curElementStageValue + " ][ " + curElementStageIndex + " ][ " + curElementStageNumbers + " ] for curElement( " + curElement.Name + " )!!!" );

		// Get the location for the new element depending on DALDirection
		let strDiagramLayout = "";
		if ( mapDiagramLayoutValues.get( DALDirection ) == DALHorizontal ) {
			strDiagramLayout = DALGetElementLocation( curElementStageIndex, curElementStageNumbers );
		} else {
			strDiagramLayout = DALGetElementLocation( curElementStageNumbers, curElementStageIndex );
		}
		BLOGDebug( "DALGetElementNewName found DALGetElementLocation[ " + curElementStageIndex + " ][ " + curElementStageNumbers + " ]: strDiagramLayout( " + strDiagramLayout + " )!!!" );

		return strDiagramLayout;
	}
	catch (catch_err)
	{
		BLOGError("DALGetElementNewName found Error: " + catch_err.message + "!!!" );
	    return "";
	}

}

/*
 * Add the Element indicated by theElementID to the diagram if it is not shown yet
 */
function DALAddElementToDiagram( theDiagram, theElementID )
{

	var curElement        as EA.Element;
	var curDiagram        as EA.Diagram;
	var curDiagramObject  as EA.DiagramObject;
	var curDiagramObjects as EA.Collection;
	var strAddNewName      = "...";

	try
	{

		// Check validity of curElement to be found in the repository by theElementID
		curElement = Repository.GetElementByID( theElementID );
		curDiagram = theDiagram;

		// Only process valid input
		if ( ( curElement == null ) || ( curDiagram == null ) ) {
			return false;
		}

		// Check all Elements in curDiagram whether the requested curElement is already shown
		curDiagramObjects = curDiagram.DiagramObjects;
		curDiagramObject  = DALGetCollectionObjectByID( curDiagramObjects, curElement.ElementID );

		// If curElement is not found on curDiagram, create a new curDiagramObject for it
		if ( curDiagramObject === null )
		{

			// Get the NewName based on the location for the new element
			strAddNewName = DALGetElementNewName( curElement );
			BLOGTrace( "DALAddElementToDiagram addNew because not found: " + curElement.Name + ", strAddNewName(" + strAddNewName + ") as part of Diagram " + curDiagram.Name + "!" );

			curDiagramObject = curDiagramObjects.AddNew( strAddNewName, "" );
			curDiagramObject.ElementID = curElement.ElementID;
			curDiagramObject.Update();
			curDiagramObjects.Refresh();

			curDiagram.Update();
			Repository.ReloadDiagram( curDiagram.DiagramID );

			return true;
		}
	}
	catch (catch_err)
	{
		BLOGError("DALAddElementToDiagram found Error: " + catch_err.message + "!!!" );
	    return false;
	}

	BLOGDebug( "DALAddElementToDiagram found " + curDiagramObject.ElementID + " as already part of curDiagram " + curDiagram.Name + "!" );
	return false;

}

/*
 * Add all elements connected to theElementID to theElementsSet
 */
function DALAddConnectedElementsToSet( theElementsSet, theElementID )
{

	// Cast the input so we get intellisense
	var curElementsSet;
	var curElement as EA.Element;
	curElementsSet  = theElementsSet;

	try
	{

		// Check validity of curElement to be found in the repository by theElementID
		curElement = Repository.GetElementByID( theElementID );
		if ( curElement == null )
		{
			BLOGWarning("DALAddConnectedElementsToSet DID NOT find theElementID(" + theElementID + ") so returns " + curElementsSet.size + " Elements as part of Diagram " + objGlobalDALDiagram.Name + "!!!" );
			return curElementsSet;
		}

		// Process all Elements connected to this Element
		var curElementConnectors as EA.Collection;
		var curConnector         as EA.Connector;
		curElementConnectors      = curElement.Connectors;

		// Check all element Connectors against data to find
		let curElementConnectorsCount = curElementConnectors.Count;
		for ( var i = 0 ; i < curElementConnectorsCount ; i++ )
		{
			curConnector = curElementConnectors.GetAt( i );

			// Check whether to include the element connected to curConnector
			if ( DALShowConnector( curConnector ) ) {

				// Find the newElementID to be on the other end of the Connector
				var newElementID = curConnector.ClientID;
				if ( curConnector.ClientID == curElement.ElementID )
				{
					newElementID = curConnector.SupplierID;
				}

				// Add the newElementID to curElementsSet
				curElementsSet.add( newElementID );
			}
			BLOGTrace("DALAddConnectedElementsToSet curElement(" + curElement.ElementID + "), newElementID=" + newElementID + ", curConnector.ClientID= " + curConnector.ClientID + ", SupplierID= " + curConnector.SupplierID + "!!!" );
		}
	}
	catch (catch_err)
	{
		BLOGError("DALAddConnectedElementsToSet found Error: " + catch_err.message + "!!!" );
	}

	// Clean up memory
	curElementConnectors = null;

	BLOGDebug("DALAddConnectedElementsToSet returns " + curElementsSet.size + " Elements as part of Diagram " + objGlobalDALDiagram.Name + "!!!" );
	return curElementsSet;
}

/*
 * Process theElementsSet for theLevel of DiagramAnalysisLayout
 */
function DALProcessNextLevel( theElementsSet, theLevel )
{

	// Only process when the maximum level not reached yet
	let curLevel = Number( theLevel );
	let maxLevel = Number( mapDiagramLayoutValues.get( DALNumLevels ) );
	if ( curLevel >= maxLevel ) {
		return;
	}
	let nextLevel = Number( theLevel ) + 1;

	try
	{

		// Start new newElementsSet for next level
		let newElementsSet   = new Set();
		let numElementsShown = 0;

		// Add elements to diagram and collect connected elements for next level
		for ( const setElementID of theElementsSet ) {

			// Add element to diagram and move pointer to next location
			if ( DALAddElementToDiagram( objGlobalDALDiagram, setElementID ) ) {
				// Got to next location
				numElementsShown++;
			}

			// Add connected elements to newElementsSet
			newElementsSet = DALAddConnectedElementsToSet( newElementsSet, setElementID );
			BLOGTrace("DALProcessNextLevel[" + theLevel + "]: setElementID = " + setElementID + ", numElementsShown = " + numElementsShown + "!" );

		}

		// Process the newElementsSet for nextLevel of DiagramAnalysisLayout if it contains elements only
		BLOGDebug("DALProcessNextLevel[" + theLevel + "]: processed " + theElementsSet.size + " Elements of which " + numElementsShown + " shown as part of Diagram " + objGlobalDALDiagram.Name + "!!!" );
		if ( newElementsSet.size > 0 ) {
			DALProcessNextLevel( newElementsSet, nextLevel );
		}
	}
	catch (catch_err)
	{
		BLOGError("DALProcessNextLevel found Error: " + catch_err.message + "!!!" );
	}

	// Clean up memory
	newElementsSet = null;

}

/*
 * Process an Element provided as parameter for DiagramAnalysisLayout
 */
function DALProcessElement( theElement )
{

	// Cast theElement to EA.Element so we get intellisense
	var curElement as EA.Element;
	curElement      = theElement;

	try
	{

		mapDALStageLevels = new Map();

		// Create the default values for mapDiagramLayoutValues
		DALCreateDiagramLayoutValues();

		// Check whether the currentElement is for DiagramAnalysisLayout
		if ( DALIsDiagramAnalysisLayoutElement( curElement ) )
		{
			BLOGTrace("DALProcessElement Process curElement(" + curElement.Name + ") for DiagramAnalysisLayout!!!" );

			// Create and get the mapDiagramLayoutValues
			DALGetValuesFromElement( curElement );

			// Get all Elements connected to this Element
			let setDiagramLayoutElements = new Set();
			setDiagramLayoutElements     = DALAddConnectedElementsToSet( setDiagramLayoutElements, curElement.ElementID );

			// Process the setDiagramLayoutElements for DiagramAnalysisLayout first Level
			DALProcessNextLevel( setDiagramLayoutElements, 0 );
			BLOGTrace("DALProcessElement curElement(" + curElement.ElementID + ") processed " + setDiagramLayoutElements.size + " Elements connected to Element(" + curElement.Name + ") for Diagram " + objGlobalDALDiagram.Name + "!!!" );

			// Clean up memory
			setDiagramLayoutElements = null;
			mapDALStageLevels        = null;

		} else {
			BLOGWarning("Element(" + curElement.Name + ") is not for DiagramAnalysisLayout!!!" );
		}
	}
	catch (catch_err)
	{
		BLOGError("DALProcessElement found Error: " + catch_err.message + "!!!" );
	}

}
