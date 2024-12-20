//[group=BaatDiagramScripts]
!INC Local Scripts.EAConstants-JavaScript
!INC BaatScriptLib.BaatScript-Logging
!INC BaatScriptLib.JavaScript-ImEx
!INC BaatScriptLib.ImportDataBricksCommon
!INC BaatScriptLib.DiagramAnalysisLayoutCommon

/*
 * This code has been included from the default Diagram Script template.
 * If you wish to modify this template, it is located in the Config\Script Templates
 * directory of your EA install path.
 *
 * Script Name:	ImportDataBricksDALGeneratie
 * Author:		J. de Baat
 * Purpose:		Generate IDB DAL Diagrams for all IDB DAL Objects found on the selected objGlobalIDBDiagram
 * Date:		20-12-2024
 *
 * This script takes a number of selected elements on a diagram to start the analysis.
 * IDBAnalyseDiagram: Analyse the DiagramObjects on this diagram
 *   If the element is a Legend then keep it for global use
 *   If the element contains the DALKeywordsTag then keep it for further processing
 * IDBProcessDALObjects: For each DALKeywordsTag element on the diagram
 *   IDBCheckOrAddPackageDiagram: Add the diagram for this element if it does not exist yet
 *   IDBCheckDALDiagram:
 *     Remove all elements from the diagram except the curElement
 *     Add the global Legend element
 *   DALProcessElement: Add the DAL elements
 * 
 */

// The level to log at
var   BLOGLEVEL = BLOGLEVEL_INFO;		// Choose from: BLOGLEVEL_ERROR, BLOGLEVEL_WARNING, BLOGLEVEL_INFO, BLOGLEVEL_DEBUG, BLOGLEVEL_TRACE

// A set of Maps to contain the different DiagramObjects found
var   arrIDBBRONObjects        = [];		// Array with IDB BRON Objects found in objGlobalIDBDiagram.DiagramObjects
var   arrIDBDALObjects         = [];		// Array with IDB DAL  Objects found in objGlobalIDBDiagram.DiagramObjects
var   objIDBDALLegendObject    = null;		// IDB DAL Legend Object       found in objGlobalIDBDiagram.DiagramObjects
var   objIDBDALLegendElement   = null;		// IDB DAL Legend Element      found in objGlobalIDBDiagram.DiagramObjects

const strAddNewNameDALObject   = "l=30;r=180;t=200;b=250;";
const strAddNewNameLegend      = "l=30;r=180;t=300;b=510;";

const strDiagramHideAttributes = "HideAtts=1;";					// Hides the Attribute compartment of Elements
const strDiagramHideConnLabel  = "HideConnStereotype=1;";		// Hides the Stereotype Labels of Connectors

/*
 * Analyse the DiagramObjects of objGlobalIDBDiagram
 */
function IDBAnalyseDiagram()
{

	var curElement        as EA.Element;
	var curDiagramObject  as EA.DiagramObject;
	var curDiagramObjects as EA.Collection;
	curDiagramObjects      = objGlobalIDBDiagram.DiagramObjects;

	try {

		// Loop over curDiagramObjects to classify the DiagramObjects
		let curDiagramObjectsCount = curDiagramObjects.Count;
		for ( var i = 0 ; i < curDiagramObjectsCount ; i++ )
		{
			curDiagramObject = curDiagramObjects.GetAt( i );

			strMessage = "( " + objGlobalIDBDiagram.Name + " ) found curDiagramObject.ElementID = " + curDiagramObject.ElementID + ",  .ObjectType = " + curDiagramObject.ObjectType + " !!!";
			BLOGTrace( strMessage );

			// Get the curElement from the curDiagramObject.ElementID found
			curElement = GetElementByID( curDiagramObject.ElementID );

			if ( curElement != null ) {
				// Check whether the currentElement is for DiagramAnalysisLayout
				if ( DALIsDiagramAnalysisLayoutElement( curElement ) ) {

					// DAL Element found
					arrIDBDALObjects.push( curElement );
					strMessage = "( " + objGlobalIDBDiagram.Name + " ) found DAL curElement.Name = " + curElement.Name + " (" + curElement.ElementID + ") for Tag = " + curElement.Tag + ", arrIDBDALObjects = " + arrIDBDALObjects.length + " !!!";
					BLOGTrace( strMessage );
				} else {
					// Check whether the currentElement is a Legend
					if ( DALElementIsLegend( curElement ) ) {

						// objIDBDALLegend Element found
						objIDBDALLegendObject  = curDiagramObject;
						objIDBDALLegendElement = curElement;
						strMessage = "( " + objGlobalIDBDiagram.Name + " ) found DAL objIDBDALLegendElement.Name = " + objIDBDALLegendElement.Name + " (" + objIDBDALLegendElement.ElementID + ") for Tag = " + objIDBDALLegendElement.Tag + ", Stereotype = " + objIDBDALLegendElement.Stereotype + " !!!";
						BLOGTrace( strMessage );
						// IDBDumpObject( objIDBDALLegendObject,  "objIDBDALLegendObject."  + objIDBDALLegendObject.ElementID );
						// IDBDumpObject( objIDBDALLegendElement, "objIDBDALLegendElement." + objIDBDALLegendElement.ElementID );
					} else {
						// curElement found is NOT Legend nor DAL Element
						strMessage = "( " + objGlobalIDBDiagram.Name + " ) found NOT Legend nor DAL Element.Name = " + curElement.Name + " (" + curElement.ElementID + ") for Tag = " + curElement.Tag + ", Stereotype = " + curElement.Stereotype + " !!!";
						BLOGTrace( strMessage );
						// IDBDumpObject( curElement, curElement.Name );
					}
				}
			}

		}

		// Show debug message
		if ( objIDBDALLegendElement != null ) {
			strMessage = "( " + objGlobalIDBDiagram.Name + " ) found DAL objIDBDALLegendElement.Name = " + objIDBDALLegendElement.Name + " (" + objIDBDALLegendElement.ElementID + ") and " + arrIDBDALObjects.length + " arrIDBDALObjects!!!";
			BLOGDebug( strMessage );
		} else {
			strMessage = "( " + objGlobalIDBDiagram.Name + " ) found NO DAL objIDBDALLegendElement and " + arrIDBDALObjects.length + " arrIDBDALObjects!!!";
			BLOGDebug( strMessage );
		}

	} catch (catch_err) {
		BLOGError( " catched error " + catch_err.message + "!" );
		return "IDBAnalyseDiagram catched error " + catch_err.message + "!";
	}

	return "";
}

/*
 * Process the arrIDBDALObjects found when analyzing objGlobalIDBDiagram
 */
function IDBProcessDALObjects()
{

	var curElement        as EA.Element;
	var curDiagramObject  as EA.DiagramObject;
	var curDiagramObjects as EA.Collection;
	var curIDBDALDiagram  as EA.Diagram;
	curDiagramObjects      = objGlobalIDBDiagram.DiagramObjects;

	try {

		// Loop over arrIDBDALObjects to process each of them
		let arrIDBDALObjectsLength = arrIDBDALObjects.length;
		for ( var i = 0 ; i < arrIDBDALObjectsLength ; i++ )
		{

			// Get the curElement from the arrIDBDALObjects array
			curElement = arrIDBDALObjects[i];

			if ( curElement != null ) {
				// Check whether there is a diagram for the curElement
				curIDBDALDiagram = IDBCheckOrAddPackageDiagram( objGlobalIDBPackage, curElement.Name );
				if ( curIDBDALDiagram == null ) {
					strMessage = "( " + curElement.Name + " ) found NO curIDBDALDiagram for " + arrIDBDALObjects.length + " arrIDBDALObjects!!!";
					BLOGError( strMessage );
					strMessage = "IDBAnalyseDiagram could not find or create curIDBDALDiagram for " + curElement.Name + " !";
					return strMessage;
				}
				strMessage = "( " + curIDBDALDiagram.Name + " ) found for curElement.Name = " + curElement.Name + " (" + i + ") of " + arrIDBDALObjects.length + " arrIDBDALObjects!!!";
				BLOGTrace( strMessage );

				// Check the Diagram for Non DAL Elements
				strMessage = IDBCheckDALDiagram( curIDBDALDiagram, curElement );

				// Process the DAL Element on the curIDBDALDiagram as objGlobalDALDiagram;
				objGlobalDALDiagram = curIDBDALDiagram;
				DALProcessElement( curElement );

				BLOGInfo( "Diagram " + curIDBDALDiagram.Name + " ( " + curIDBDALDiagram.DiagramID + " ) is processed at " + _BLOGGetDisplayDate() + "!" );
			}

		}

	} catch (catch_err) {
		BLOGError( " catched error " + catch_err.message + "!" );
		return "IDBAnalyseDiagram catched error " + catch_err.message + "!";
	}

	return "";
}

/*
 * Process the arrIDBDALObjects found when analyzing objGlobalIDBDiagram
 */
function IDBCheckDALDiagram( theDiagram, theElement )
{

	// Cast the input values to objects so we get intellisense
	var curElement        as EA.Element;
	var curDiagram        as EA.Diagram;
	var curDiagramObject  as EA.DiagramObject;
	var curDiagramObjects as EA.Collection;
	var curIDBDALDiagram  as EA.Diagram;

	let boolDALElementFound = false;

	curElement = theElement;
	curDiagram = theDiagram;

	// Only process valid input
	if ( ( curElement != null ) && ( curDiagram != null ) ) {

		curDiagramObjects = curDiagram.DiagramObjects;

		// Loop over curDiagramObjects to delete all except curElement
		let curDiagramObjectsCount = curDiagramObjects.Count;
		for ( var i = curDiagramObjectsCount - 1 ; i >= 0 ; i-- )
		{

			// Get the curElement from the arrIDBDALObjects array
			curDiagramObject = curDiagramObjects.GetAt( i );

			if ( curDiagramObject != null ) {
				// Check whether the curDiagramObject is curElement
				if ( curDiagramObject.ElementID === curElement.ElementID ) {

					boolDALElementFound = true;

					// Skip curElement found
					strMessage = "( " + curDiagram.Name + " ) SKIP DAL curElement.Name = " + curElement.Name + " (" + curElement.ElementID + ") for curDiagramObject: " + curDiagramObject.ElementID + "!!!";
					BLOGTrace( strMessage );
				} else {
					// Remove the curDiagramObject from the diagram
					curDiagramObjects.DeleteAt( i, false );
					strMessage = "( " + curDiagram.Name + " ) DeleteAt Non DAL curDiagramObject: " + curDiagramObject.ElementID + "!!!";
					BLOGTrace( strMessage );
				}
			}

		}

		// Show progress
		curDiagramObjects.Refresh();
		strMessage = "( " + curDiagram.Name + " ) Deleted Non DAL objects except curElement.Name = " + curElement.Name + " (" + curElement.ElementID + ") from curDiagramObjectsCount: " + curDiagramObjectsCount + " to " + curDiagramObjects.Count + "!!!";
		BLOGDebug( strMessage );

		// Add the curElement as DAL Object to curDiagram if not found
		if ( boolDALElementFound === false ) {
			curDiagramObject = curDiagramObjects.AddNew( strAddNewNameDALObject, "" );
			curDiagramObject.ElementID = curElement.ElementID;
			curDiagramObject.Update();
			curDiagramObjects.Refresh();
			strMessage = "( " + curDiagram.Name + " ) Added curElement.Name = " + curElement.Name + " (" + curElement.ElementID + ") with strAddNewNameDALObject: " + strAddNewNameDALObject + "!!!";
			BLOGTrace( strMessage );
		}

		// Add the objIDBDALLegendObject to curDiagram
		curDiagramObject = curDiagramObjects.AddNew( strAddNewNameLegend, "" );
		curDiagramObject.ElementID = objIDBDALLegendElement.ElementID;
		curDiagramObject.Update();
		curDiagramObjects.Refresh();

		strMessage = "( " + curDiagram.Name + " ) Added objIDBDALLegendElement.Name = " + objIDBDALLegendElement.Name + " (" + objIDBDALLegendElement.ElementID + ") with strAddNewNameLegend: " + strAddNewNameLegend + "!!!";
		BLOGTrace( strMessage );
		curDiagram.Update();

		// Set the properties of curDiagram
		curDiagram.HighlightImports = false;		// Hides the Namespace of Elements
		curDiagram.ExtendedStyle    = strDiagramHideAttributes;
		curDiagram.StyleEx          = strDiagramHideConnLabel;
		curDiagram.Update();

		Repository.ReloadDiagram( curDiagram.DiagramID );

	}

	return "";
}

/*
 * Dump a parsed JSON object
 */
function IDBDumpObject( theObject, theObjectName )
{

	try
	{

		BLOGTrace(" Found Attributes of theObject.Name ( " + theObjectName + " )" );
		for ( let key in theObject ) {
			BLOGTrace( "==> theObject.[" + key + "]= " + theObject[key] + " !" );
		}
	}
	catch(catch_err)
	{
		BLOGError("IDBDumpObject found Error: " + catch_err + "!!!" );
		return null;
	}

}

/*
 * Diagram Script main function
 */
function ImportDataBricksDALGeneratie()
{
	// Show the script output window
	Repository.EnsureOutputVisible( "Script" );

	Session.Output( "======================================= Started ImportDataBricksDALGeneratie at " + _BLOGGetDisplayDate() + "!" );

	let curResult = "";

	try
	{

		// Get and check the global variables
		const validDiagram = IDBGetAndCheckDiagram();
		if ( validDiagram ) {

			// Analyse the DiagramObjects of objGlobalIDBDiagram
			if ( curResult.length == 0 ) {
				BLOGInfo( "Diagram( " + objGlobalIDBDiagram.Name + " ) is VALID so proceed processing!" );
				curResult = IDBAnalyseDiagram();
			}

			// Process the arrIDBDALObjects found when analyzing objGlobalIDBDiagram
			if ( curResult.length == 0 ) {
				BLOGInfo( "Diagram( " + objGlobalIDBDiagram.Name + " ) is ANALYSED so proceed processing!" );
				curResult = IDBProcessDALObjects();
			}

			// Handle curResult or Finish processing
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

	Session.Output( "======================================= Finished ImportDataBricksDALGeneratie at " + _BLOGGetDisplayDate() + "!" );
}

ImportDataBricksDALGeneratie();
