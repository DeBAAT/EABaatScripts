//[group=BaatDiagramScripts]
!INC Local Scripts.EAConstants-JavaScript
!INC EAScriptLib.JavaScript-Logging
!INC EAScriptLib.JavaScript-Dialog

/*
 * Script Name:	DiagramAnalysisLayout
 * Author:		J de Baat
 * Purpose:		Dynamically draw elements on the diagram as indicated by the TaggedValues defined in the selected element
 * Date:		25-08-2024
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

/*
 * mapDiagramLayoutValues is a map of [Key,Value] to use when generating the analysis diagram
 * Each mapDiagramLayoutValue has a default defined which can be replaced with a value defined as TaggedValue for the selected element
 */
var mapDiagramLayoutValues;

var theGlobalDiagram as EA.Diagram;

const DALDirection     = "DALDirection";			// Whether to analyse in Horizontal (default) or Vertical Direction
const DALNumLevels     = "DALNumLevels";			// Number of levels to analyse
const DALSpaceHor      = "DALSpaceHor";				// Horizontal spacing between elements
const DALSpaceVer      = "DALSpaceVer";				// Vertical spacing between elements
const DALStartHor      = "DALStartHor";				// Horizontal start for elements
const DALStartVer      = "DALStartVer";				// Vertical start for elements
const DALElementWidth  = "DALElementWidth";			// Width of an element to draw
const DALElementHeight = "DALElementHeight";		// Height of an element to draw

const DALTagPrefix     = "DAL";						// String prefix to indicate that this TaggedValue can be used for DiagramAnalysisLayout
const DALKeywordsTag   = "diagramanalysislayout";	// String Tag to indicate that this element can be used for DiagramAnalysisLayout
const DALHorizontal    = "Horizontal";				// Analyse in Horizontal (default) Direction
const DALVertical      = "Vertical";				// Analyse in Vertical Direction

/*
 * Create the set of defaults for mapDiagramLayoutValues
 */
function createDiagramLayoutValues()
{

	// Clean up memory and create a new Map
	mapDiagramLayoutValues = null;
	mapDiagramLayoutValues = new Map();

	mapDiagramLayoutValues.set( DALDirection,     DALHorizontal );
	mapDiagramLayoutValues.set( DALNumLevels,     4             );
	mapDiagramLayoutValues.set( DALSpaceHor,      20            );
	mapDiagramLayoutValues.set( DALSpaceVer,      20            );
	mapDiagramLayoutValues.set( DALStartHor,      50            );
	mapDiagramLayoutValues.set( DALStartVer,      100           );
	mapDiagramLayoutValues.set( DALElementWidth,  150           );
	mapDiagramLayoutValues.set( DALElementHeight, 60            );

}

/*
 * Get the set of values for mapDiagramLayoutValues from the TaggedValues of the Element
 */
function getDiagramLayoutValuesFromElement( theElement )
{

	// Cast the input values to objects so we get intellisense
	var curElement      as EA.Element;
	var curElementTags  as EA.Collection;
	var curTaggedValue  as EA.TaggedValue;

	curElement           = theElement;
	curElementTags       = curElement.TaggedValues;

	// Check all element tags for mapDiagramLayoutValues (starting with "DAL")
	let curElementTagsCount = curElementTags.Count;
	for ( var i = 0 ; i < curElementTagsCount ; i++ )
	{
		curTaggedValue = curElementTags.GetAt( i );
		if ( (curTaggedValue.Name.startsWith( DALTagPrefix )) ) 
		{
			// Get the mapDiagramLayoutValues from the curTaggedValue
			mapDiagramLayoutValues.set( curTaggedValue.Name,   curTaggedValue.Value );
			// Session.Output("Found curTaggedValue to use for mapDiagramLayoutValues[ " + curTaggedValue.Name + " ] = " + curTaggedValue.Value + " !!!" );
		}
	}

	// Clean up memory
	curElementTags = null;

	return false;

}

/*
 * Get the location of an element based on the parameters
 */
function getElementLocation( numX, numY )
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

/*
 * An element is for DiagramAnalysisLayout when the KeywordsTag contains the string defined by DALKeywordsTag
 */
function isDiagramAnalysisLayoutElement( theElement )
{
	// Cast theElement to EA.Element so we get intellisense
	var inputElement as EA.Element;
	inputElement      = theElement;

	try
	{
		var strKeywordsTag = inputElement.Tag.toLowerCase();
		var idxKeywordsTag = strKeywordsTag.indexOf( DALKeywordsTag );
		return ( idxKeywordsTag >= 0 );
	}
	catch(e)
	{
	    return false;
	}

}

/*
 * Get the requested Element from theCollectionObjects using theElementID as parameter
 */
function getCollectionObjectByID( theCollectionObjects, theElementID )
{
	// Cast theElement to EA.Element so we get intellisense
	var curCollectionObjects as EA.Collection;
	var curCollectionObject  as EA.Element;

	// Check all Elements in theCollectionObjects whether the requested curElement is already defined
	curCollectionObjects = theCollectionObjects;

	// Loop over curCollectionObjects to find theElementID
	let curCollectionObjectsCount = curCollectionObjects.Count;
	for ( var i = 0 ; i < curCollectionObjectsCount ; i++ )
	{
		curCollectionObject = curCollectionObjects.GetAt( i );
		if ( curCollectionObject.ElementID === theElementID ) {
			// Session.Output( "getCollectionObjectByID found CollectionObject ( " + curCollectionObject.ElementID + " ) as part of " + curCollectionObjectsCount + " CollectionObjects!!!" );
			return curCollectionObject;
		}
		// Session.Output( "getCollectionObjectByID TESTED CollectionObject ( " + curCollectionObject.ElementID + " ) against theElementID " + theElementID + " !!!" );
	}

	// theElementName not found as part of curCollectionObjects
	// Session.Output( "getCollectionObjectByID DID NOT FIND theElementID ( " + theElementID + " ) as part of " + curCollectionObjectsCount + " CollectionObjects!!!" );
	return null;

}

/*
 * Add the Element indicated by theElementID to the diagram if it is not shown yet
 */
function AddElementToDiagram( theElementID, numX, numY )
{

	var curElement        as EA.Element;
	var curDiagramObjects as EA.Collection;
	var curDiagramObject  as EA.DiagramObject;
	var strAddNewName      = "...";

	// Check validity of curElement to be found in the repository by theElementID
	curElement = Repository.GetElementByID( theElementID );
	if ( curElement == null )
	{
		return false;
	}

	// Check all Elements in theDiagram whether the requested curElement is already shown
	curDiagramObjects = theGlobalDiagram.DiagramObjects;
	curDiagramObject  = getCollectionObjectByID( curDiagramObjects, curElement.ElementID );

	// If curElement is not found on theGlobalDiagram, create a new curDiagramObject for it
	if ( curDiagramObject === null )
	{

		// Get the location for the new element
		strAddNewName = getElementLocation( numX, numY );
		// Session.Output( "AddElementToDiagram addNew because not found: " + curElement.Name + ", strAddNewName(" + strAddNewName + ") as part of Diagram " + theGlobalDiagram.Name + ", pos(" + numX + "," + numY + ") !" );

		curDiagramObject = curDiagramObjects.AddNew( strAddNewName, "" );
		curDiagramObject.ElementID = curElement.ElementID;
		curDiagramObject.Update();
		curDiagramObjects.Refresh();

		theGlobalDiagram.Update();
		Repository.ReloadDiagram( theGlobalDiagram.DiagramID );

		return true;
	}

	// Session.Output( "AddElementToDiagram found " + curDiagramObject.ElementID + " as part of Diagram " + theGlobalDiagram.Name + ", pos(" + numX + "," + numY + ") !" );
	return false;

}

/*
 * Add all elements connected to theElementID to theElementsSet
 */
function AddConnectedElementsToSet( theElementsSet, theElementID )
{

	// Cast the input so we get intellisense
	var curElementsSet;
	var curElement as EA.Element;
	curElementsSet  = theElementsSet;

	// Check validity of curElement to be found in the repository by theElementID
	curElement = Repository.GetElementByID( theElementID );
	if ( curElement == null )
	{
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

		// Find the newElementID to be on the other end of the Connector
		var newElementID = curConnector.ClientID;
		if ( curConnector.ClientID == curElement.ElementID )
		{
			newElementID = curConnector.SupplierID;
		}

		// Add the newElementID to curElementsSet
		curElementsSet.add( newElementID );
		// Session.Output("AddConnectedElementsToSet curElement(" + curElement.ElementID + "), newElementID=" + newElementID + ", curConnector.ClientID= " + curConnector.ClientID + ", SupplierID= " + curConnector.SupplierID + "!!!" );
	}

	// Clean up memory
	curElementConnectors = null;

	// Session.Output("AddConnectedElementsToSet returns " + curElementsSet.size + " Elements as part of Diagram " + theGlobalDiagram.Name + "!!!" );
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

	// Get first location for theLevel depending on DALDirection
	var curHor = 0;
	var curVer = 0;
	if ( mapDiagramLayoutValues.get( DALDirection ) == DALHorizontal ) {
		curHor = theLevel;
		curVer = 0;
	} else {
		curHor = 0;
		curVer = theLevel;
	}

	// Start new newElementsSet for next level
	let newElementsSet   = new Set();
	let numElementsShown = 0;

	// Add elements to diagram and collect connected elements for next level
	for ( const setElementID of theElementsSet ) {

		// Add element to diagram and move pointer to next location
		if ( AddElementToDiagram( setElementID, curHor, curVer ) ) {
			// Got to next location
			numElementsShown++;
			if ( mapDiagramLayoutValues.get( DALDirection ) == DALHorizontal ) {
				curVer++;
			} else {
				curHor++;
			}

			// Add connected elements to newElementsSet
			newElementsSet = AddConnectedElementsToSet( newElementsSet, setElementID );
		}

		// Session.Output("DALProcessNextLevel[" + theLevel + "]: setElementID = " + setElementID + ", Location = [" + curHor + "," + curVer + "] !" );

	}

	// Session.Output("DALProcessNextLevel[" + theLevel + "]: processed " + theElementsSet.size + " Elements as part of Diagram " + theGlobalDiagram.Name + "!!!" );

	// Process the newElementsSet for nextLevel of DiagramAnalysisLayout if it contains elements only
	Session.Output("DALProcessNextLevel[" + theLevel + "]: processed " + theElementsSet.size + " Elements of which " + numElementsShown + " shown as part of Diagram " + theGlobalDiagram.Name + "!!!" );
	if ( newElementsSet.size > 0 ) {
		DALProcessNextLevel( newElementsSet, nextLevel );
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

	// Create the default values for mapDiagramLayoutValues
	createDiagramLayoutValues();

	// Check whether the currentElement is for DiagramAnalysisLayout
	if ( isDiagramAnalysisLayoutElement( curElement ) )
	{
		// Session.Output("DALProcessElement Process curElement(" + curElement.Name + ") for DiagramAnalysisLayout!!!" );

		// Create and get the mapDiagramLayoutValues
		getDiagramLayoutValuesFromElement( curElement );

		// Get all Elements connected to this Element
		let setDiagramLayoutElements = new Set();
		setDiagramLayoutElements     = AddConnectedElementsToSet( setDiagramLayoutElements, curElement.ElementID );

		// Process the setDiagramLayoutElements for DiagramAnalysisLayout first Level
		DALProcessNextLevel( setDiagramLayoutElements, 0 );
		Session.Output("DALProcessElement curElement(" + curElement.ElementID + ") processed " + setDiagramLayoutElements.size + " Elements connected to Element(" + curElement.Name + ") for Diagram " + theGlobalDiagram.Name + "!!!" );

		// Clean up memory
		setDiagramLayoutElements = null;

	} else {
		Session.Output("Element(" + curElement.Name + ") is not for DiagramAnalysisLayout!!!" );
	}

}

/*
 * Diagram Script main function
 */
function DiagramAnalysisLayout()
{
	// Show the script output window
	Repository.EnsureOutputVisible( "Script" );

	Session.Output( "======================================= Started DiagramAnalysisLayout " );

	// Get a reference to theGlobalDiagram
	theGlobalDiagram = Repository.GetCurrentDiagram();

	if ( theGlobalDiagram != null )
	{
		// Get a reference to any selected objects
		var selectedElement as EA.Element;
		var selectedObjects as EA.Collection;
		selectedObjects      = theGlobalDiagram.SelectedObjects;

		if ( selectedObjects.Count > 0 )
		{
			Session.Output("Selected selectedObjects.Count: " + selectedObjects.Count );

			// One or more diagram objects are selected
			let selectedObjectsCount = selectedObjects.Count;
			for ( var i = 0 ; i < selectedObjectsCount ; i++ )
			{
				// Process the currentDiagramElement
				var currentDiagramElement as EA.Element;
				var currentElement        as EA.Element;
				currentDiagramElement      = selectedObjects.GetAt( i );
				currentElement             = Repository.GetElementByID( currentDiagramElement.ElementID );

				// Process the currentElement for DiagramAnalysisLayout
				DALProcessElement( currentElement );
			}

			// Reload theGlobalDiagram when all processing is done
			Repository.ReloadDiagram( theGlobalDiagram.DiagramID );

		}
		else
		{
			// Nothing is selected
			LOGError( "This script requires at least one element to be selected!" );
		}

		Session.Output( "======================================= Finished DiagramAnalysisLayout " );
	}
	else
	{
		Session.Prompt( "This script requires a diagram to be visible.", promptOK);
	}
}

DiagramAnalysisLayout();
